/* ============================================================
   Code.gs — Aplicación de Reservas (Google Apps Script)
   ============================================================

   CONFIGURACIÓN DE LA HOJA:
   ─────────────────────────
   Pestaña "Users":
     Col A = Nombre
     Col B = Email
     Col C = PIN (numérico, ej. 1234)
     Fila 1 = encabezado, datos desde fila 2

   Pestaña "Calendars":
     Col A = Calendar ID
     Col B = Nombre de la sala
     Col C = Color hex (ej. #3B82F6)
     Col D = Centro (ej. "El Poblet" o "La Sala")
     Col E = Grupo de combinación ("1" para salas combinables, vacío si no)
     Col F = URL de Google Maps (ej. https://maps.app.goo.gl/...)
     Fila 1 = encabezado, datos desde fila 2

   DESPLIEGUE:
   ───────────
   Implementar → Nueva implementación → Aplicación web
     - Ejecutar como: Yo
     - Acceso: Cualquier persona
   ============================================================ */


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
// 2. SESSION MANAGEMENT (CacheService — no OAuth needed)
// ─────────────────────────────────────────────────────────────
var SESSION_TTL = 21600; // 6 hours

function generateToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var t = '';
  for (var i = 0; i < 48; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

function saveSession(token, email) {
  CacheService.getScriptCache().put('sess_' + token, email, SESSION_TTL);
}

function getSessionEmail(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get('sess_' + token) || null;
}

function deleteSession(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
}

function requireAuth(token) {
  var email = getSessionEmail(token);
  if (!email) throw new Error('SESSION_EXPIRED');
  return email;
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

/** Users tab: A=Name, B=Email, C=PIN */
function getUsers() {
  var sheet = getSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error('No se encontró la pestaña "Users".');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(function(r) { return r[1] && r[1].toString().trim(); })
    .map(function(r) {
      return {
        name:  r[0] ? r[0].toString().trim() : '',
        email: r[1].toString().toLowerCase().trim(),
        pin:   r[2] ? r[2].toString().trim() : ''
      };
    });
}

/** Calendars tab: A=ID, B=Name, C=Color, D=Centre, E=PairGroup */
function getCalendarConfig() {
  var sheet = getSpreadsheet().getSheetByName('Calendars');
  if (!sheet) throw new Error('No se encontró la pestaña "Calendars".');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    .filter(function(r) { return r[0] && r[0].toString().trim(); })
    .map(function(r) {
      return {
        id:        r[0].toString().trim(),
        name:      r[1] ? r[1].toString().trim() : 'Sala sin nombre',
        color:     (r[2] ? r[2].toString().trim() : '') || '#3B82F6',
        centre:    r[3] ? r[3].toString().trim() : '',
        pairGroup: r[4] ? r[4].toString().trim() : '',
        mapsUrl:   r[5] ? r[5].toString().trim() : ''
      };
    });
}

function getDistinctCentres(calendars) {
  var seen = {}, out = [];
  calendars.forEach(function(c) {
    if (c.centre && !seen[c.centre]) { seen[c.centre] = true; out.push(c.centre); }
  });
  return out;
}


// ─────────────────────────────────────────────────────────────
// 4. AUTH ENDPOINTS
// ─────────────────────────────────────────────────────────────

function login(email, pin) {
  try {
    email = email.toLowerCase().trim();
    pin   = pin.toString().trim();

    var users = getUsers();
    var user  = users.find(function(u) { return u.email === email; });

    if (!user)     return { success: false, message: 'Email no reconocido. Contacta al administrador.' };
    if (!user.pin) return { success: false, message: 'Sin PIN asignado. Contacta al administrador.' };
    if (user.pin !== pin) return { success: false, message: 'PIN incorrecto. Inténtalo de nuevo.' };

    var token     = generateToken();
    saveSession(token, email);

    var calendars = getCalendarConfig();
    var safeUsers = users.map(function(u) { return { name: u.name, email: u.email }; });

    return {
      success:   true,
      token:     token,
      email:     email,
      userName:  user.name,
      users:     safeUsers,
      calendars: calendars,
      centres:   getDistinctCentres(calendars)
    };
  } catch (err) {
    Logger.log('login error: ' + err);
    return { success: false, message: 'Error del servidor: ' + err.message };
  }
}

function resumeSession(token) {
  try {
    var email = getSessionEmail(token);
    if (!email) return { success: false };

    var users = getUsers();
    var user  = users.find(function(u) { return u.email === email; });
    if (!user) return { success: false };

    var calendars = getCalendarConfig();
    var safeUsers = users.map(function(u) { return { name: u.name, email: u.email }; });

    return {
      success:   true,
      token:     token,
      email:     email,
      userName:  user.name,
      users:     safeUsers,
      calendars: calendars,
      centres:   getDistinctCentres(calendars)
    };
  } catch (err) {
    return { success: false };
  }
}

function logout(token) {
  deleteSession(token);
}


// ─────────────────────────────────────────────────────────────
// 5. FETCH CALENDAR EVENTS
// ─────────────────────────────────────────────────────────────

function getCalendarEvents(token, calendarId, startStr, endStr) {
  requireAuth(token);
  var cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('Calendario no encontrado: ' + calendarId);

  return cal.getEvents(new Date(startStr), new Date(endStr)).map(function(e) {
    var isRec = false;
    try { isRec = e.isRecurringEvent(); } catch (x) {}
    var guests = [];
    try {
      guests = e.getGuestList(true).map(function(g) {
        return { email: g.getEmail(), name: g.getName() || g.getEmail(), status: g.getGuestStatus().toString() };
      });
    } catch (x) {}
    return {
      id: e.getId(), title: e.getTitle(),
      start: e.getStartTime().toISOString(), end: e.getEndTime().toISOString(),
      description: e.getDescription() || '', location: e.getLocation() || '',
      guests: guests, isRecurring: isRec, color: isRec ? '#9CA3AF' : null
    };
  });
}

function getMultiCalendarEvents(token, calendarIds, startStr, endStr) {
  requireAuth(token);
  var all = [];
  calendarIds.forEach(function(calId) {
    try {
      var cal = CalendarApp.getCalendarById(calId);
      if (!cal) return;
      cal.getEvents(new Date(startStr), new Date(endStr)).forEach(function(e) {
        var isRec = false;
        try { isRec = e.isRecurringEvent(); } catch (x) {}
        all.push({
          id: e.getId(), title: e.getTitle(),
          start: e.getStartTime().toISOString(), end: e.getEndTime().toISOString(),
          description: e.getDescription() || '', location: e.getLocation() || '',
          guests: [], isRecurring: isRec, color: isRec ? '#9CA3AF' : null,
          sourceCalendarId: calId
        });
      });
    } catch (err) { Logger.log('Error fetching ' + calId + ': ' + err); }
  });
  all.sort(function(a, b) { return new Date(a.start) - new Date(b.start); });
  return all;
}


// ─────────────────────────────────────────────────────────────
// 6. CREATE BOOKING
// ─────────────────────────────────────────────────────────────

function createBooking(token, formObject) {
  var bookerEmail = requireAuth(token);

  if (!formObject.calendarIds || !formObject.startTime || !formObject.duration || !formObject.title) {
    throw new Error('Faltan campos obligatorios.');
  }

  var startTime = new Date(formObject.startTime);
  var duration  = parseInt(formObject.duration, 10);
  if (isNaN(duration) || duration <= 0) throw new Error('Duración no válida.');
  var endTime = new Date(startTime.getTime() + duration * 60000);

  var cals = formObject.calendarIds.map(function(id) {
    var cal = CalendarApp.getCalendarById(id);
    if (!cal) throw new Error('Calendario no encontrado: ' + id);
    return cal;
  });

  // Conflict check
  var conflicts = [];
  cals.forEach(function(cal) {
    var ex = cal.getEvents(startTime, endTime);
    if (ex.length) conflicts.push(cal.getName() + ': "' + ex[0].getTitle() + '"');
  });
  if (conflicts.length) throw new Error('¡Conflicto de horario!\n' + conflicts.join('\n'));

  // Guests
  var gs = {};
  [bookerEmail, formObject.organiser, formObject.responsible, formObject.open, formObject.close]
    .forEach(function(em) { if (em && em.indexOf('@') !== -1) gs[em.toLowerCase().trim()] = true; });

  // Description
  var roles = [
    '────────────────────────',
    'ROLES:',
    '  Organizador:  ' + (formObject.organiser  || '–'),
    '  Tel. Organiz: ' + (formObject.phone       || '–'),
    '  Responsable:  ' + (formObject.responsible || '–'),
    '  Apertura:     ' + (formObject.open        || '–'),
    '  Cierre:       ' + (formObject.close       || '–'),
    '────────────────────────',
    'Reservado por: ' + bookerEmail
  ].join('\n');

  var desc = formObject.description
    ? formObject.description.trim() + '\n\n' + roles
    : roles;

  var ids = cals.map(function(cal) {
    return cal.createEvent(formObject.title, startTime, endTime, {
      description: desc,
      guests:      Object.keys(gs).join(','),
      sendInvites: true,
      location:    formObject.location || ''
    }).getId();
  });

  return {
    status:   'success',
    eventIds: ids,
    message:  'Reserva confirmada para ' +
              startTime.toLocaleDateString('es-ES') + ' a las ' +
              startTime.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' })
  };
}


// ─────────────────────────────────────────────────────────────
// 7. DEBUG HELPERS
// ─────────────────────────────────────────────────────────────
function debugSheet() {
  Logger.log('USERS: ' + JSON.stringify(getUsers()));
  Logger.log('CALENDARS: ' + JSON.stringify(getCalendarConfig()));
}
function debugCalendarAccess() {
  getCalendarConfig().forEach(function(c) {
    var cal = CalendarApp.getCalendarById(c.id);
    Logger.log((cal ? '✔' : '✗') + ' ' + c.name + ' [' + c.centre + '] ' + c.id);
  });
}
