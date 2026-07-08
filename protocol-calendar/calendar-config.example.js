window.PROTOCOL_CALENDAR_CONFIG = {
  // Register a Microsoft Entra ID single-page application and paste its client ID here.
  // Redirect URI should match the URL used to open this page.
  // Leave blank to keep Outlook sync disabled; saved marks can still be downloaded as one .ics file.
  microsoftClientId: "",
  authority: "https://login.microsoftonline.com/common",
  scopes: ["Calendars.ReadWrite", "MailboxSettings.ReadWrite"],
};
