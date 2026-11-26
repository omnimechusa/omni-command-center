// Omni Command Center (phone app) settings. No secrets live here: every read
// goes through the signed-in person's own Microsoft 365 login.
window.OMNI_CONFIG = {
  // Microsoft Entra app registration "Omni Command Center" (single tenant, SPA).
  // Leave blank until it exists; the app then shows a one-time setup screen.
  clientId: "14c75dde-8270-473c-a47a-14094eedd3e2",
  tenant: "omnimechusa.com",

  // OmniMech SharePoint "Documents" library
  driveId: "b!VCLIe5VBy0mC-a2NHeazsJQaiZS1FcVOtZr6D9ctzCESTSZyOy4oTYdaKpuoncaz",
  paths: {
    boards: "Open Construction Projects/Claude outputs/boards",
    jobBoardPublish: "Accounting/Tom/JobBoard/publish",
    permits: "Open Construction Projects/Claude outputs/permits/desk-data.json",
    toolAccounts: "Accounting/AP/Employee Tool Purchase Assistance Program/Tool_Account_Master.xlsx",
    uniformOrders: "Open Construction Projects/Claude outputs/boards/marketing/uniform/orders.json"
  },

  // Mail sources
  joyce: "Joyce.Aguilar@omnimechusa.com",
  sigma: "info@send.sigmacomputing.com",
  sigmaReports: ["Daily Reports Weekly", "Weekly Timecard Exceptions Report"],

  // Microsoft To Do lists shown on Command (names; matched case-insensitive)
  todoMain: "Tom Current To Do",

  // Links that open in their own app/window
  links: {
    uniformApp: "https://claude.ai/artifact/DmjEGWzYQApfNVmiU3Tfws",
    claudeBoard: "https://claude.ai/artifact/Ni1J4DsD6J2it5XdBC1Epi",
    buildops: "https://live.buildops.com",
    toolPolicy: "https://claude.ai/artifact/BGDxrkfQhrKSmLzbrAfXTs"
  }
};
