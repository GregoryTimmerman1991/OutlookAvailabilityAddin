// outlook_availability_addin.js
// Custom Outlook Add-in to insert weekly availability (UK and CET) in a styled table

Office.onReady(() => {
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);
});

function formatTimeSlot(start, end, timeZone) {
  return `${start.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', timeZone })} - ` +
         `${end.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', timeZone })}`;
}

async function insertAvailabilityTable() {
  try {
    const weeks = prompt("Enter number of weeks to include (e.g., 1 or 2):", "1");
    if (!weeks || isNaN(weeks) || weeks < 1) return;

    const weekStartInput = prompt("Enter start date of the week (YYYY-MM-DD):");
    const startOfWeek = new Date(weekStartInput);
    if (isNaN(startOfWeek.getTime())) return;

    const calendarItems = await Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async result => {
      if (result.status !== "succeeded") return;

      const token = result.value;
      const mailbox = Office.context.mailbox;

      let allRows = "";

      for (let w = 0; w < weeks; w++) {
        const weekStart = new Date(startOfWeek);
        weekStart.setDate(weekStart.getDate() + w * 7);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);

        const startISO = weekStart.toISOString();
        const endISO = weekEnd.toISOString();
        const restUrl = `${mailbox.restUrl}/v2.0/me/calendarview?startDateTime=${startISO}&endDateTime=${endISO}`;

        const response = await fetch(restUrl, {
          method: "GET",
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json"
          }
        });

        const data = await response.json();

        const tableRows = data.value.map(event => {
          const start = new Date(event.Start.DateTime);
          const end = new Date(event.End.DateTime);
          const ukHour = start.toLocaleTimeString('en-GB', { hour: '2-digit', hour12: false, timeZone: "Europe/London" });
          if (ukHour === "09") return ""; // Skip 9:00-10:00 UK

          const ukTime = formatTimeSlot(start, end, "Europe/London");
          const cetTime = formatTimeSlot(start, end, "Europe/Brussels");

          return `
            <tr>
              <td>${start.toLocaleDateString("en-GB", { weekday: 'long', day: 'numeric', month: 'short' })}</td>
              <td>${ukTime}</td>
              <td>${cetTime}</td>
            </tr>
          `;
        }).join("\n");

        allRows += `
          <p><b>Gregory's Availability – Week of ${weekStart.toDateString()}</b></p>
          <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
            <thead>
              <tr style="background-color: #003A2B; color: white;">
                <th style="padding: 8px; border: 1px solid #003A2B;">Day</th>
                <th style="padding: 8px; border: 1px solid #003A2B;">UK Time</th>
                <th style="padding: 8px; border: 1px solid #003A2B;">CET Time</th>
              </tr>
            </thead>
            <tbody>
              ${tableRows}
            </tbody>
          </table>
        `;
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(allRows, { coercionType: Office.CoercionType.Html });
    });
  } catch (error) {
    console.error("Error inserting availability:", error);
  }
}

window.insertAvailabilityTable = insertAvailabilityTable;
