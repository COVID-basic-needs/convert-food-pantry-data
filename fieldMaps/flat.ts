export default {
  siteName: "Organization Name",
  siteStreetAddress: "Street Address",
  siteCity: "City",
  siteState: "State",
  siteZip: "Zip",
  contactPhone: "General Phone",
  contactEmail: "Organization Email",
  siteType: () => [],
  siteCountry: () => "USA",
  siteSubType: () => [],
  "Site Needs/Updates Forms": () => [],
  Claims: () => [],
  url: "Website",
  "Notes (possibly Pre-COVID)": (fields) =>
    `${fields["General Services Details"]}\n${fields["Notes"]}`,
};