window.PROTOCOL_CALENDAR_CONFIG = {
  // Register a Microsoft Entra ID single-page application and paste its client ID here.
  // Redirect URI should match the URL used to open this page.
  // Leave blank to use the built-in .ics download, which imports into Google Calendar or Outlook.
  microsoftClientId: "",
  authority: "https://login.microsoftonline.com/common",
  scopes: ["Calendars.ReadWrite"],
};
