import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

// MSAL instance — lazily initialized by initMsal()
let msalInstance = undefined;

/**
 * Initialize MSAL as a nestable public client application.
 * Called lazily before each token acquisition so onReady stays fast.
 */
async function initMsal() {
  if (!msalInstance) {
    const msalConfig = {
      auth: {
        clientId: "202ce61a-109e-4291-a605-733fbaf6f77f",
        authority: "https://login.microsoftonline.com/common",
      },
      cache: {
        cacheLocation: "localStorage",
      },
    };
    msalInstance = await createNestablePublicClientApplication(msalConfig);
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Check NAA is supported before wiring the button
    if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
      showStatus("Nested App Auth is not supported in this Outlook version.", true);
      return;
    }
    document.getElementById("helloButton").onclick = sayHello;
  }
});

function showStatus(message, isError) {
  const el = document.getElementById("status");
  el.style.color = isError ? "red" : "green";
  el.textContent = message;
}

/**
 * Acquire an access token: try silent first, fall back to popup if interaction required.
 */
async function getToken() {
  await initMsal();

  const tokenRequest = { scopes: ["User.Read"] };

  let accessToken = null;

  try {
    const userAccount = await msalInstance.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (silentError) {
    if (silentError instanceof InteractionRequiredAuthError) {
      console.log(`Unable to acquire token silently: ${silentError}`);
      // Silent acquisition failed — fall back to interactive popup
      try {
        const userAccount = await msalInstance.acquireTokenPopup(tokenRequest);
        console.log("Acquired token interactively.");
        accessToken = userAccount.accessToken;
      } catch (popupError) {
        console.error(`Unable to acquire token interactively: ${popupError}`);
        throw popupError;
      }
    } else {
      // Error cannot be resolved through interaction
      console.error(`Unable to acquire token silently: ${silentError}`);
      throw silentError;
    }
  }

  return accessToken;
}

/**
 * Fetches user profile from Microsoft Graph and inserts signature into email body.
 */
async function sayHello() {
  showStatus("", false);

  let token;
  try {
    token = await getToken();
    showStatus("Token obtained successfully.", false);
  } catch (error) {
    showStatus("Failed to get token: " + error.message, true);
    return;
  }

  let me;
  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,jobTitle,department,companyName,mobilePhone,businessPhones,city,state,country,postalCode,streetAddress",
      { headers: { Authorization: "Bearer " + token } }
    );

    if (!response.ok) {
      showStatus("Graph API error: " + response.status + " " + response.statusText, true);
      return;
    }

    me = await response.json();
  } catch (error) {
    showStatus("Graph API request failed: " + error.message, true);
    return;
  }

  const phone = me.mobilePhone || (me.businessPhones && me.businessPhones[0]) || "";

  const signature = `<span style="COLOR: black; FONT-FAMILY: 
  'Calibri'; FONT-SIZE: 12pt;"><strong>${me.displayName || ""}
  </strong></span><span style="FONT-SIZE: 11pt; COLOR: black; FONT-FAMILY: 'Calibri'">
  <br>${me.jobTitle || ""}<br>
  </span><span style="FONT-SIZE: 10pt; COLOR: gray; FONT-FAMILY: 'Calibri'">
  <p style="margin:0">
  ${me.companyName ? me.companyName + "<br>" : ""}
  ${me.department ? me.department + "<br>" : ""}
  ${me.streetAddress ? me.streetAddress + ", ": ""} 
  ${me.postalCode ? me.postalCode + ", ": ""}
  ${me.city ? me.city + ", ": ""}
  ${me.country ? me.country + "<br><br>" : ""} 
  ${"telefon: +420 " + me.mobilePhone ? me.mobilePhone + "<br>" : ""}
  ${"e-mail: " + me.mail ? "e-mail: " + me.mail : ""}
  </p></span>
  <br><a href="http://www.ssc.cas.cz/cs/">
  <img alt="" height="40" src="https://cms11.avcr.cz/export/sites/stredisko-spolecnych-cinnosti/.content/galerie-obrazku/podpisy_obraz_zdroje/ssc_logo-01.png" width="235">
  </a><br><br>
  <table border="0" cellpadding="0" cellspacing="0" style="margin-top:6px;margin-bottom:6px;">
    <tr>
      <td style="padding-right:16px;">
        <a href="https://www.ssc.cas.cz" style="text-decoration:none;">
          <img alt="" height="16" src="https://cms11.avcr.cz/export/sites/stredisko-spolecnych-cinnosti/.content/galerie-obrazku/podpisy_obraz_zdroje/ikona_web.png" width="16" style="vertical-align:middle;">
          <span style="FONT-SIZE: 10pt; COLOR: gray; FONT-FAMILY: 'Calibri'; vertical-align:middle;">&nbsp;Web SSČ AV ČR</span>
        </a>
      </td>
      <td style="padding-right:16px;">
        <a href="https://www.ssc.cas.cz/nabidka-sluzeb" style="text-decoration:none;">
          <img alt="" height="16" src="https://cms11.avcr.cz/export/sites/stredisko-spolecnych-cinnosti/.content/galerie-obrazku/podpisy_obraz_zdroje/ikona_brozura.png" width="16" style="vertical-align:middle;">
          <span style="FONT-SIZE: 10pt; COLOR: gray; FONT-FAMILY: 'Calibri'; vertical-align:middle;">&nbsp;Nabídka služeb</span>
        </a>
      </td>
      <td>
        <a href="https://www.linkedin.com/company/sscavcr" style="text-decoration:none;">
          <img alt="" height="16" src="https://cms11.avcr.cz/export/sites/stredisko-spolecnych-cinnosti/.content/galerie-obrazku/podpisy_obraz_zdroje/ikona_linkedin.png" width="16" style="vertical-align:middle;">
          <span style="FONT-SIZE: 10pt; COLOR: gray; FONT-FAMILY: 'Calibri'; vertical-align:middle;">&nbsp;LinkedIn</span>
        </a>
      </td>
    </tr>
  </table>
  <br><a href="http://www.ssc.cas.cz/cs/banner">
  <img alt="" height="84" src="http://www.ssc.cas.cz/cs/banner/banner.jpg" width="400">
  </a><a href="https://www.ssc.cas.cz/cs/kariera/HR-Award/">
  <table width="400" border="0" cellpadding="0">
  <tr>
  <td width="80" align="left"><img alt="" height="46" src="https://cms11.avcr.cz/export/sites/stredisko-spolecnych-cinnosti/.content/galerie-obrazku/podpisy_obraz_zdroje/hr_award.png" width="68">
  </td><td width="320" align="left"><span style="FONT-SIZE: 10pt; COLOR: gray; FONT-FAMILY: 'Calibri'">Jsme držitelem HR Excellence in Research Award</span></td>
  </tr>
  </table></a>`;

  Office.context.mailbox.item.body.setSignatureAsync(
    signature,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showStatus("Failed to insert signature: " + asyncResult.error.message, true);
      }
    }
  );
}
