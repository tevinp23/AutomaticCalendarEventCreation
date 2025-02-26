/*************  ✨ Codeium Command ⭐  *************/
/**
 * Creates calendar events for existing responses on the "Form Responses 1" sheet.
 * Looks for the following columns:
 *   - Event Name (6)
 *   - Event Date (8)
 *   - Setup Time (9)
 *   - Start Time (10)
 *   - End Time (11)
 * If any of these columns are empty for a row, will log a message and skip it.
 * If a row is missing a required field, will log a message and skip it.
 * Logs a message for each row successfully processed with the event name, start time, and end time.
 * Logs a message for each row with an error.
 * Logs a message when done processing all responses.
 */
/******  0f98f493-25b9-4318-9595-3e7882e0c4f7  *******/
function createEventsForExistingResponses() {
    const timeZone = "America/New_York"; // Eastern Time
    const calendarId = "id"; 
    const calendar = CalendarApp.getCalendarById(calendarId);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
    
    if (!sheet) {
      throw new Error("Sheet named 'Form Responses 1' not found!");
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("No responses to process.");
      return;
    }
    
    // Helper function to parse a "hh:mm:ss AM/PM" time string into hours, minutes, seconds.
    function parseTimeString(timeStr) {
      // If timeStr isn't a string, convert it (in case it's a Date or other type).
      if (typeof timeStr !== "string") {
        timeStr = Utilities.formatDate(new Date(timeStr), timeZone, "hh:mm:ss a");
      }
      timeStr = timeStr.trim();
      
      // Expect the format "hh:mm:ss AM/PM"
      const parts = timeStr.split(" ");
      if (parts.length !== 2) {
        throw new Error("Invalid time format (expected 'hh:mm:ss AM/PM'): " + timeStr);
      }
      
      const timePart = parts[0];
      const meridiem = parts[1].toUpperCase();
      const timeComponents = timePart.split(":");
      if (timeComponents.length !== 3) {
        throw new Error("Time must include hours, minutes, and seconds: " + timeStr);
      }
      
      let hours = parseInt(timeComponents[0], 10);
      const minutes = parseInt(timeComponents[1], 10);
      const seconds = parseInt(timeComponents[2], 10);
      
      if (meridiem === "PM" && hours < 12) {
        hours += 12;
      }
      if (meridiem === "AM" && hours === 12) {
        hours = 0;
      }
      return { hours, minutes, seconds };
    }
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
     
  
      const eventName    = row[6];
      const eventDate    = row[8];
      const setUpTimeStr = row[9];
      const startTimeStr = row[10];
      const endTimeStr   = row[11];
      
      if (!eventName || !eventDate || !setUpTimeStr || !startTimeStr || !endTimeStr) {
        Logger.log(`Skipping row ${i+1}: missing required fields.`);
        continue;
      }
      
      let baseDate = eventDate instanceof Date ? new Date(eventDate) : new Date(eventDate);
      baseDate.setHours(0, 0, 0, 0);
      
      const startDate = new Date(baseDate);
      const endDate   = new Date(baseDate);
      try {
        const startTime = parseTimeString(startTimeStr);
        const endTime = parseTimeString(endTimeStr);
        
        startDate.setHours(startTime.hours, startTime.minutes, startTime.seconds, 0);
        endDate.setHours(endTime.hours, endTime.minutes, endTime.seconds, 0);
      } catch (error) {
        Logger.log(`Error parsing time on row ${i+1}: ${error.message}`);
        continue;
      }
      
       const descriptionLines = [
        "Company/Department Name: " + row[1],
        "First Name: " + row[2],
        "Last Name: " + row[3],
        "Primary Phone Number: " + row[4],
        "Primary Email: " + row[5],
        "Event Name: " + row[6],
        "Event Category: " + row[7],
        "Event Date: " + row[8],
        "Room Reservation: " + row[12],
        "Alternative Dates" + row[13],
        "Est. Attendance: " + row[14],
        "Orgs/Departments Involved? " + row[15],
        "On Site Point of Contact: " + row[16],
        "Event Summary: " + row[17],
        "Audience: " + row[18],
        "Elected Official? " + row[19],
        "Candidate for Public Office? " + row[20],
        "Campaign on behalf of someone? " + row[22],
        "Speaker/Presenter with National/International Recognition? " + row[23],
        "Preferred Event Space: " + row[24],
        "AV/Equipment Needed: " + row[25],
        "Additional Details: " + row[26]
      ];
      const fullDescription = descriptionLines.join("\n");
      
      Logger.log(`Row ${i+1} -> Event: ${eventName}, Start: ${Utilities.formatDate(startDate, timeZone, "MM/dd/yyyy hh:mm:ss a")}, End: ${Utilities.formatDate(endDate, timeZone, "MM/dd/yyyy hh:mm:ss a")}`);
      
      try {
        calendar.createEvent(eventName, startDate, endDate, { description: fullDescription });
        Logger.log(`Event created for row ${i+1}: ${eventName}`);
      } catch (error) {
        Logger.log(`Error creating event for row ${i+1}: ${error.message}`);
      }
    }
    
    Logger.log("Done processing all responses.");
  }