/* ============================================================
   Code.gs — Aplicación de Reservas de Calendario (Google Apps Script)
   ============================================================
   
   INSTRUCCIONES DE DESPLIEGUE:
   ============================
   1. Crear una hoja de Google con dos pestañas:
   
      "Users" → Col A: Nombre, Col B: Email, Col C: PIN
      (fila 1 = encabezado, datos desde fila 2)
      Ejemplo:
        | Nombre      | Email              | PIN  |
        | Ana García  | ana@ejemplo.com    | 1234 |
        | Pedro López | pedro@ejemplo.com  | 5678 |
   
      "Calendars" → Col A: Calendar ID, Col B: Nombre corto, Col C: Color hex
      (fila 1 = encabezado, datos desde fila 2)
   
   2. Abrir Extensiones → Apps Script, pegar Code.gs e Index.html.
   
   3. Implementar → Nueva implementación → Aplicación web:
      - Ejecutar como: **Yo** (tu cuenta)
      - Quién tiene acceso: **Cualquier persona**
      - Hacer clic en Implementar, autorizar cuando se solicite.
   
   4. Compartir la URL de la aplicación web con los usuarios aprobados.
   
   CÓMO FUNCIONA LA AUTENTICACIÓN:
   ================================
   - El script se ejecuta con TU cuenta → acceso a hojas y calendarios
   - Los visitantes introducen su email + PIN (definido en la hoja Users)
   - Se crea un token de sesión almacenado en CacheService (6 horas)
   - Cada llamada al servidor incluye el token para verificar la sesión
   - No se requiere OAuth ni permisos de Google para los visitantes
   ============================================================ */


// ─────────────────────────────────────────────────────────────
// 0. CONFIGURATION
// ─────────────────────────────────────────────────────────────

var SESSION_DURATION_SECONDS = 6 * 60 * 60; // 6 hours


// ─────────────────────────────────────────────────────────────
// 1. SERVE THE WEB APP
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Reservas de Calendario')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ─────────────────────────────────────────────────────────────
// 2. SESSION MANAGEMENT
// ─────────────────────────────────────────────────────────────

/**
 * Generates a random session token.
 */
function generateToken_() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 48; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

/**
 * Stores a session: token → email mapping in ScriptCache.
 */
function createSession_(email) {
  var token = generateToken_();
  var cache = CacheService.getScriptCache();
  cache.put('session_' + token, email, SESSION_DURATION_SECONDS);
  return token;
}

/**
 * Validates a session token and returns the associated email.
 * Returns empty string if invalid/expired.
 */
function getSessionEmail_(token) {
  if (!token) return '';
  var cache = CacheService.getScriptCache();
  var email = cache.get('session_' + token);
  return email || '';
}

/**
 * Removes a session from cache.
 */
function destroySession_(token) {
  if (!token) return;
  var cache = CacheService.getScriptCache();
  cache.remove('session_' + token);
}


// ─────────────────────────────────────────────────────────────
// 3. SPREADSHEET HELPERS
// ─────────────────────────────────────────────────────────────

function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error('El script debe estar vinculado a una hoja de Google.');
  }
}

/**
 * Reads the "Users" tab.
 * Expects: Col A = Name, Col B = Email, Col C = PIN
 * (header in row 1, data from row 2).
 */
function getApprovedUsers_() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) throw new Error('No se encontró la pestaña "Users".');

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return data
    .filter(function(row) { return row[1] && row[1].toString().trim() !== ''; })
    .map(function(row) {
      return {
        name:  row[0] ? row[0].toString().trim() : '',
        email: row[1].toString().toLowerCase().trim(),
        pin:   row[2] ? row[2].toString().trim() : ''
      };
    });
}

/**
 * Returns user list WITHOUT pins (safe to send to client).
 */
function getApprovedUsersPublic_() {
  return getApprovedUsers_().map(function(u) {
    return { name: u.name, email: u.email };
  });
}

/**
 * Reads the "Calendars" tab and returns an array of {id, name, color}.
 */
function getCalendarConfig() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Calendars');
  if (!sheet) throw new Error('No se encontró la pestaña "Calendars".');

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return data
    .filter(function(row) { return row[0] && row[0].toString().trim() !== ''; })
    .map(function(row) {
      return {
        id:    row[0].toString().trim(),
        name:  row[1] ? row[1].toString().trim() : 'Calendario sin nombre',
        color: (row[2] ? row[2].toString().trim() : '') || '#3B82F6'
      };
    });
}


// ─────────────────────────────────────────────────────────────
// 4. LOGIN / LOGOUT
// ─────────────────────────────────────────────────────────────

/**
 * Authenticates a user with email + PIN.
 * Returns { success, token, email, userName, ... } or { success: false, message }.
 */
function login(email, pin) {
  if (!email || !pin) {
    return { success: false, message: 'Introduce tu email y PIN.' };
  }

  email = email.toLowerCase().trim();
  pin = pin.toString().trim();

  var users = getApprovedUsers_();
  var match = users.find(function(u) { return u.email === email; });

  if (!match) {
    return { success: false, message: 'Email no encontrado en la lista de usuarios.' };
  }

  if (match.pin !== pin) {
    return { success: false, message: 'PIN incorrecto.' };
  }

  var token = createSession_(email);
  var calendars = getCalendarConfig();

  return {
    success:   true,
    token:     token,
    email:     email,
    userName:  match.name,
    users:     getApprovedUsersPublic_(),
    calendars: calendars
  };
}

/**
 * Validates an existing session token and returns app data if valid.
 */
function resumeSession(token) {
  var email = getSessionEmail_(token);
  if (!email) {
    return { success: false, message: 'Sesión expirada. Inicia sesión de nuevo.' };
  }

  var users = getApprovedUsers_();
  var match = users.find(function(u) { return u.email === email; });

  if (!match) {
    destroySession_(token);
    return { success: false, message: 'Tu cuenta ha sido eliminada de la lista de usuarios.' };
  }

  var calendars = getCalendarConfig();

  return {
    success:   true,
    token:     token,
    email:     email,
    userName:  match.name,
    users:     getApprovedUsersPublic_(),
    calendars: calendars
  };
}

/**
 * Logs the user out.
 */
function logout(token) {
  destroySession_(token);
  return { success: true };
}


// ─────────────────────────────────────────────────────────────
// 5. ACCESS CHECK HELPER
// ─────────────────────────────────────────────────────────────

/**
 * Validates a session token and ensures the user is still approved.
 * Returns { email, userName } or throws.
 */
function requireSession_(token) {
  var email = getSessionEmail_(token);
  if (!email) {
    throw new Error('SESSION_EXPIRED');
  }

  var users = getApprovedUsers_();
  var match = users.find(function(u) { return u.email === email; });

  if (!match) {
    destroySession_(token);
    throw new Error('SESSION_EXPIRED');
  }

  return { email: email, userName: match.name };
}


// ─────────────────────────────────────────────────────────────
// 6. FETCH CALENDAR EVENTS
// ─────────────────────────────────────────────────────────────

function getCalendarEvents(token, calendarId, startStr, endStr) {
  requireSession_(token);

  var cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) {
    throw new Error(
      'Calendario no encontrado o acceso denegado para ID: ' + calendarId +
      '. Asegúrate de que el calendario esté compartido con la cuenta del propietario del script.'
    );
  }

  var start = new Date(startStr);
  var end   = new Date(endStr);
  var events = cal.getEvents(start, end);

  return events.map(function(e) {
    var isRecurring = false;
    try { isRecurring = e.isRecurringEvent(); } catch (err) {}

    var guestList = [];
    try {
      var guests = e.getGuestList(true);
      guestList = guests.map(function(g) {
        return {
          email: g.getEmail(),
          name:  g.getName() || g.getEmail(),
          status: g.getGuestStatus().toString()
        };
      });
    } catch (err) {}

    return {
      id:          e.getId(),
      title:       e.getTitle() + (isRecurring ? ' 🔁' : ''),
      start:       e.getStartTime().toISOString(),
      end:         e.getEndTime().toISOString(),
      description: e.getDescription() || '',
      guests:      guestList,
      isRecurring: isRecurring,
      color:       isRecurring ? '#9CA3AF' : null,
      location:    e.getLocation() || ''
    };
  });
}


// ─────────────────────────────────────────────────────────────
// 7. CREATE A BOOKING
// ─────────────────────────────────────────────────────────────

function createBooking(token, formObject) {
  var user = requireSession_(token);
  var bookerEmail = user.email;

  if (!formObject.calendarId || !formObject.startTime || !formObject.duration || !formObject.title) {
    throw new Error('Faltan campos obligatorios.');
  }

  var cal = CalendarApp.getCalendarById(formObject.calendarId);
  if (!cal) throw new Error('Calendario no encontrado o acceso denegado.');

  var startTime = new Date(formObject.startTime);
  var durationMinutes = parseInt(formObject.duration, 10);

  if (isNaN(durationMinutes) || durationMinutes <= 0) {
    throw new Error('Duración no válida.');
  }

  var endTime = new Date(startTime.getTime() + durationMinutes * 60000);

  var conflicts = cal.getEvents(startTime, endTime);
  if (conflicts.length > 0) {
    throw new Error(
      '¡Conflicto de horario! Ya existe un evento de ' +
      conflicts[0].getStartTime().toLocaleTimeString() + ' a ' +
      conflicts[0].getEndTime().toLocaleTimeString() + '. Elige otro horario.'
    );
  }

  var guestSet = {};
  function addGuest(email) {
    if (email && email.indexOf('@') !== -1) {
      guestSet[email.toLowerCase().trim()] = true;
    }
  }

  addGuest(bookerEmail);
  addGuest(formObject.organiser);
  addGuest(formObject.responsible);
  addGuest(formObject.open);
  addGuest(formObject.close);

  var guestString = Object.keys(guestSet).join(',');

  var roleBlock = [
    '────────────────────────',
    'ROLES:',
    '  Organizador:  ' + (formObject.organiser    || '—'),
    '  Responsable:  ' + (formObject.responsible   || '—'),
    '  Apertura:     ' + (formObject.open          || '—'),
    '  Cierre:       ' + (formObject.close         || '—'),
    '────────────────────────',
    'Reservado por: ' + bookerEmail
  ].join('\n');

  var fullDescription = '';
  if (formObject.description && formObject.description.trim()) {
    fullDescription = formObject.description.trim() + '\n\n' + roleBlock;
  } else {
    fullDescription = roleBlock;
  }

  var event = cal.createEvent(formObject.title, startTime, endTime, {
    description: fullDescription,
    guests:      guestString,
    sendInvites: true
  });

  return {
    status: 'success',
    eventId: event.getId(),
    message: 'Reserva confirmada para ' + startTime.toLocaleDateString('es-ES') +
             ' a las ' + startTime.toLocaleTimeString('es-ES', {hour: '2-digit', minute:'2-digit'})
  };
}


// ─────────────────────────────────────────────────────────────
// 8. DEBUG HELPERS
// ─────────────────────────────────────────────────────────────

function debugConfig() {
  Logger.log('=== USUARIOS APROBADOS ===');
  var users = getApprovedUsers_();
  users.forEach(function(u) {
    Logger.log('  ' + u.name + ' | ' + u.email + ' | PIN: ' + (u.pin ? '****' : 'NO PIN'));
  });
  Logger.log('');
  Logger.log('=== CALENDARIOS ===');
  Logger.log(JSON.stringify(getCalendarConfig()));
}

function debugCalendarAccess() {
  var calendars = getCalendarConfig();
  calendars.forEach(function(c) {
    var cal = CalendarApp.getCalendarById(c.id);
    Logger.log((cal ? '✓' : '✗') + ' ' + c.name + ' (' + c.id + ')');
  });
}
