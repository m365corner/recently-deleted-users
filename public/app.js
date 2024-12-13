const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<client-id-goes-here>",
        authority: "https://login.microsoftonline.com/<tenant-id-goes-here>",
        redirectUri: "http://localhost:8000",
    },
});

let allDeletedUsers = [];

// Login
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["Directory.Read.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");
        await fetchDeletedUsers();
    } catch (error) {
        console.error("Login failed:", error);
        alert("Login failed.");
    }
}

function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Fetch Recently Deleted Users
async function fetchDeletedUsers() {
    const currentDate = new Date();
    const startDate = new Date(currentDate.setDate(currentDate.getDate() - 180)).toISOString();

    const response = await callGraphApi(`/directory/deletedItems/microsoft.graph.user?$filter=deletedDateTime ge ${startDate}&$select=displayName,userPrincipalName,mail,assignedLicenses,deletedDateTime`);
    allDeletedUsers = response.value;
}

// Search Function
function search() {
    const searchText = document.getElementById("searchBox").value.toLowerCase();
    const fromDate = document.getElementById("fromDate").value ? new Date(document.getElementById("fromDate").value).toISOString() : null;
    const toDate = document.getElementById("toDate").value ? new Date(document.getElementById("toDate").value).toISOString() : null;
    

    const filteredUsers = allDeletedUsers.filter(user => {
        const matchesSearchText = searchText
            ? (user.displayName?.toLowerCase().includes(searchText) ||
               user.userPrincipalName?.toLowerCase().includes(searchText) ||
               user.mail?.toLowerCase().includes(searchText))
            : true;

        const matchesDateRange = fromDate && toDate
            ? new Date(user.deletedDateTime) >= new Date(fromDate) && new Date(user.deletedDateTime) <= new Date(toDate)
            : true;

            

        return matchesSearchText && matchesDateRange; //&& matchesLicense;
    });

    if (filteredUsers.length === 0) {
        alert("No matching results found.");
    }

    displayResults(filteredUsers);
}

// Display Results
function displayResults(users) {
    const outputBody = document.getElementById("outputBody");
    outputBody.innerHTML = users.map(user => `
        <tr>
            <td>${user.displayName || "N/A"}</td>
            <td>${user.userPrincipalName || "N/A"}</td>
            <td>${user.mail || "N/A"}</td>
            <td>${user.assignedLicenses.length > 0 ? "Licensed" : "Unlicensed"}</td>
            <td>${new Date(user.deletedDateTime).toLocaleDateString()}</td>
        </tr>
    `).join("");
}

// Utility Function to Call Graph API
async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Directory.Read.All", "Mail.Send"],
        account,
    });

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method,
        headers: {
            Authorization: `Bearer ${tokenResponse.accessToken}`,
            "Content-Type": "application/json",
        },
        body: body ? JSON.stringify(body) : null,
    });

    if (response.ok) {
        const contentType = response.headers.get("content-type");
        if (contentType && contentType.includes("application/json")) {
            return await response.json();
        }
        return {};
    } else {
        const error = await response.text();
        console.error(`Graph API Error (${response.status}):`, error);
        throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
    }
}

// Download Report as CSV
function downloadReportAsCSV() {
    const headers = ["Display Name", "UPN", "Email", "License Status", "Deleted Date"];
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data available to download.");
        return;
    }

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Recently_Deleted_Users_Report.csv";
    link.click();
}

// Mail Report to Admin
async function sendReportAsMail() {
    const adminEmail = document.getElementById("adminEmail").value;

    if (!adminEmail) {
        alert("Please provide an admin email.");
        return;
    }

    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data to send via email.");
        return;
    }

    const emailContent = rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
    const emailBody = `
        <table border="1">
            <thead>
                <tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>${emailContent}</tbody>
        </table>
    `;

    const message = {
        message: {
            subject: "Recently Deleted Users Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: adminEmail } }],
        },
    };

    try {
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent successfully!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report.");
    }
}
