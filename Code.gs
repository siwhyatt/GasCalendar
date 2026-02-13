/* ============================================================
   Code.gs — Aplicación de Reservas de Calendario (Google Apps Script)
   ============================================================
   
   INSTRUCCIONES DE DESPLIEGUE:
   ============================
   1. Crear una hoja de Google con dos pestañas:
      - "Users"     → Col A: Nombre, Col B: Email (fila 1 = encabezado, datos desde fila 2)
      - "Calendars" → Col A: Calendar ID, Col B: Nombre corto, Col C: Color hex (#3B82F6)
   
   2. Abrir Extensiones → Apps Script, pegar Code.gs e Index.html.
   
   3. En el editor de Apps Script, ir a Configuración del proyecto (⚙️):
      - Marcar "Mostrar archivo de manifiesto appsscript.json en el editor"
      - Editar appsscript.json con los oauthScopes listados abajo.
   
   4. Implementar → Nueva implementación → Aplicación web:
      - Ejecutar como: **Yo** (tu cuenta de Workspace)
      - Quién tiene acceso: **Cualquier persona con cuenta de Google**
      - Hacer clic en Implementar, autorizar cuando se solicite.
   
   5. Compartir la URL de la aplicación web con los usuarios aprobados.
   
   REQUIRED OAUTH SCOPES (add to appsscript.json):
   ================================================
   {
     "timeZone": "Europe/London",
     "dependencies": {},
     "webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_LOGIN" },
     "exceptionLogging": "STACKDRIVER",
     "oauthScopes": [
       "https://www.googleapis.com/auth/spreadsheets.readonly",
       "https://www.googleapis.com/auth/calendar",
       "https://www.googleapis.com/auth/calendar.events",
       "https://www.googleapis.com/auth/userinfo.email",
       "https://www.googleapis.com/auth/script.external_request"
     ]
   }
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
// 2. ROBUST EMAIL DETECTION
// ─────────────────────────────────────────────────────────────

function getUserEmail() {
  var email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (err) {
    Logger.log('getActiveUser failed: ' + err);
  }
  if (email) return email.toLowerCase().trim();

  try {
    email = Session.getEffectiveUser().getEmail();
  } catch (err) {
    Logger.log('getEffectiveUser failed: ' + err);
  }
  if (email) return email.toLowerCase().trim();

  try {
    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v1/userinfo?alt=json', {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var json = JSON.parse(resp.getContentText());
    if (json.email) return json.email.toLowerCase().trim();
  } catch (err) {
    Logger.log('OAuth userinfo failed: ' + err);
  }

  return '';
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
 * Expects: Col A = Name, Col B = Email (header in row 1, data from row 2).
 * Returns an array of { name: "...", email: "..." } objects.
 */
function getApprovedUsers() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) throw new Error('No se encontró la pestaña "Users".');

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return data
    .filter(function(row) { return row[1] && row[1].toString().trim() !== ''; })
    .map(function(row) {
      return {
        name:  row[0] ? row[0].toString().trim() : '',
        email: row[1].toString().toLowerCase().trim()
      };
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
  Logger.log('Calendars tab lastRow: ' + lastRow);

  if (lastRow < 2) {
    Logger.log('La pestaña Calendars no tiene filas de datos.');
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  Logger.log('Raw calendar data: ' + JSON.stringify(data));

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
// 4. MAIN DATA ENDPOINT
// ─────────────────────────────────────────────────────────────

function getAppData() {
  var email = getUserEmail();
  var userList = getApprovedUsers();

  Logger.log('Detected email: "' + email + '"');
  Logger.log('Approved users: ' + JSON.stringify(userList));

  if (!email) {
    return {
      hasAccess: false,
      email: '',
      error: 'NO_EMAIL_DETECTED',
      message: 'No se pudo detectar tu cuenta de Google. Asegúrate de haber iniciado sesión e inténtalo de nuevo.'
    };
  }

  var emailList = userList.map(function(u) { return u.email; });
  if (emailList.indexOf(email) === -1) {
    return {
      hasAccess: false,
      email: email,
      error: 'NOT_APPROVED',
      message: 'Tu email (' + email + ') no está en la lista de usuarios aprobados. Contacta al administrador.'
    };
  }

  var currentUser = userList.find(function(u) { return u.email === email; });
  var calendars = getCalendarConfig();

  return {
    hasAccess: true,
    email: email,
    userName: currentUser ? currentUser.name : '',
    users: userList,        // [{name, email}, ...] for role dropdowns
    calendars: calendars
  };
}


// ─────────────────────────────────────────────────────────────
// 5. FETCH CALENDAR EVENTS
// ─────────────────────────────────────────────────────────────

function getCalendarEvents(calendarId, startStr, endStr) {
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
    } catch (err) {
      Logger.log('No se pudo obtener la lista de invitados: ' + err);
    }

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
// 6. CREATE A BOOKING
// ─────────────────────────────────────────────────────────────

function createBooking(formObject) {
  if (!formObject.calendarId || !formObject.startTime || !formObject.duration || !formObject.title) {
    throw new Error('Faltan campos obligatorios. Por favor completa calendario, hora, duración y título.');
  }

  var cal = CalendarApp.getCalendarById(formObject.calendarId);
  if (!cal) throw new Error('Calendario no encontrado o acceso denegado.');

  var bookerEmail = getUserEmail();
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
// 7. DEBUG / TEST HELPERS
// ─────────────────────────────────────────────────────────────

function debugAuth() {
  Logger.log('=== AUTH DEBUG ===');
  Logger.log('getActiveUser:    "' + Session.getActiveUser().getEmail() + '"');
  Logger.log('getEffectiveUser: "' + Session.getEffectiveUser().getEmail() + '"');
  Logger.log('getUserEmail():   "' + getUserEmail() + '"');
  Logger.log('');
  Logger.log('=== USUARIOS APROBADOS ===');
  Logger.log(JSON.stringify(getApprovedUsers()));
  Logger.log('');
  Logger.log('=== CALENDARIOS ===');
  Logger.log(JSON.stringify(getCalendarConfig()));
}

function debugCalendarAccess() {
  var calendars = getCalendarConfig();
  calendars.forEach(function(c) {
    var cal = CalendarApp.getCalendarById(c.id);
    if (cal) {
      Logger.log('✓ Acceso OK: ' + c.name + ' (' + c.id + ')');
    } else {
      Logger.log('✗ SIN ACCESO: ' + c.name + ' (' + c.id + ')');
    }
  });
}

function debugCalendarsTab() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Calendars');
  if (!sheet) { Logger.log('ERROR: No se encontró pestaña "Calendars"'); return; }

  Logger.log('Sheet name: ' + sheet.getName());
  Logger.log('getLastRow(): ' + sheet.getLastRow());
  Logger.log('getLastColumn(): ' + sheet.getLastColumn());

  var range = sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), 10), 3);
  var values = range.getValues();
  values.forEach(function(row, i) {
    Logger.log('Row ' + (i + 1) + ': A="' + row[0] + '" | B="' + row[1] + '" | C="' + row[2] + '"');
  });

  var result = getCalendarConfig();
  Logger.log('getCalendarConfig() returned: ' + JSON.stringify(result));
}
