import "./outlook";

// Initialisierung des Add-Ins für Outlook
Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        // Office-Initialisierungsfunktion
        info.office.context.mailbox.addHandlerAsync(
            Office.EventType.AppointmentAttendeeCommandSurfaceChanged,
            onAppointmentAttendeeCommandSurfaceChanged
        );
    }
});

// Reaktion auf Änderungen im AppointmentAttendeeCommandSurface
async function onAppointmentAttendeeCommandSurfaceChanged(event) {
    const surface = event.surface;
    if (surface === Office.MailboxEnums.AppointmentAttendeeCommandSurfaceLocation.ComposeAppointment) {
        // Benutzer erstellt einen Termin
        const isOutOfOffice = await checkIfOutOfOffice();
        if (isOutOfOffice) {
            //Rufe Start- und Enddatum aus dem Terminformular
            const{start, end}=getAppointmentDates();

            saveAutoReplySettings(autoReplyMessage, startDate, endDate);
            // Hier kannst du auch benutzerdefinierten Code für die automatische Antwort hinzufügen
            showOutOfOfficeDialog();
        }
    }
}

// Funktion zum Abrufen von Start- und Enddatum aus dem Terminformular
function getAppointmentDates() {
    const appointmentItem = Office.context.mailbox.item;
    
    // Hier gehst du davon aus, dass start und end im ISO-Format (YYYY-MM-DDTHH:mm:ss.sssZ) vorliegen
    const start = new Date(appointmentItem.start);
    const end = new Date(appointmentItem.end);

    return { start, end };
}

// Überprüfung, ob der Benutzer außer Haus ist
async function checkIfOutOfOffice() {
    const appointmentItem = Office.context.mailbox.item;

    // Annahme: Der Wert "Out of Office" wird durch "Private" in der sensitivity-Eigenschaft dargestellt
    const sensitivity = appointmentItem.sensitivity;

    // Überprüfe, ob die sensitivity-Eigenschaft auf den gewünschten Status gesetzt ist
    const isOutOfOffice = sensitivity && sensitivity === "Private";

    return isOutOfOffice;
}

async function inspectAppointmentProperties() {
    const appointmentItem = Office.context.mailbox.item;
    console.log(appointmentItem);
}



// Funktion zum Speichern der automatischen Antwort-Einstellungen
function saveAutoReplySettings(message, start, end) {
    // Implementiere Logik zum Speichern der Einstellungen (z.B., in den Office-Add-In-Eigenschaften)
    console.log("Einstellungen gespeichert:", message, start, end);
}

// Anzeige des Dialogfensters
function showOutOfOfficeDialog() {
    // Hier sollte die Logik zur Anzeige des Dialogfensters implementiert werden

}
