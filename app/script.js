async function setBusy() {
  const token = await getToken();
  const response = await fetch("https://graph.microsoft.com/v1.0/me/presence/setPresence", {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      sessionId: "status-keeper-app",
      availability: "Busy",
      activity: "InACall",
      expirationDuration: "PT10M"
    })
  });
  if (response.ok) {
    alert("Statut défini à 'Occupé' pour 10 minutes.");
  } else {
    alert("Erreur lors du changement de statut.");
  }
}

async function getToken() {
  const msalConfig = {
    auth: {
      clientId: "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
      authority: "https://login.microsoftonline.com/common"
    }
  };
  const msalInstance = new msal.PublicClientApplication(msalConfig);
  const result = await msalInstance.ssoSilent({ scopes: ["Presence.ReadWrite"] });
  return result.accessToken;
}
