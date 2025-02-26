function createEventOnFormSubmit(e) {
  const timeZone = "America/New_York"; 
  const calendarId = "id"; 
  const calendar = CalendarApp.getCalendarById(calendarId);
  

  const row = e.values;
  
  const eventName    = row[6];
  const eventDate    = row[8];
  const setUpTimeStr = row[9];
  const startTimeStr = row[10];
  const endTimeStr   = row[11];
  
  if (!eventName || !eventDate || !setUpTimeStr || !startTimeStr || !endTimeStr) {
    Logger.log("Missing required fields in the form submission. No event created.");
    return;
  }
  
  // Helper function to parse a "hh:mm:ss AM/PM" time string.
  function parseTimeString(timeStr) {
    timeStr = timeStr.trim(); // Remove any leading/trailing spaces
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
    // Convert to 24-hour time.
    if (meridiem === "PM" && hours < 12) {
      hours += 12;
    }
    if (meridiem === "AM" && hours === 12) {
      hours = 0;
    }
    return { hours, minutes, seconds };
  }
  
  // Reset the event date to midnight
  let baseDate = eventDate instanceof Date ? new Date(eventDate) : new Date(eventDate);
  baseDate.setHours(0, 0, 0, 0);
  
  // Compute start and end Date objects.
  const startDate = new Date(baseDate);
  const endDate = new Date(baseDate);
  try {
    const startTime = parseTimeString(startTimeStr);
    const endTime = parseTimeString(endTimeStr);
    
    startDate.setHours(startTime.hours, startTime.minutes, startTime.seconds, 0);
    endDate.setHours(endTime.hours, endTime.minutes, endTime.seconds, 0);
  } catch (error) {
    Logger.log("Error parsing time: " + error.message);
    return;
  }
  
  // Build the event description using your original format.
  const descriptionLines = [
    "Set-up Time: " + row[9],
    "Company/Department Name: " + row[1],
    "First Name: " + row[2],
    "Last Name: " + row[3],
    "Primary Phone Number: " + row[4],
    "Primary Email: " + row[5],
    "Event Name: " + row[6],
    "Event Category: " + row[7],
    "Event Date: " + row[8],
    "Alternative Dates: " + row[12],
    "Est. Attendance: " + row[14],
    "Orgs/Departments Involved?: " + row[15],
    "On Site Point of Contact: " + row[16],
    "Event Summary: " + row[17],
    "Audience: " + row[18],
    "Elected Official?: " + row[19],
    "Candidate for Public Office?: " + row[20],
    "Campaign on behalf of someone?: " + row[22],
    "Speaker/Presenter with National/International Recognition?: " + row[23],
    "Preferred Event Space: " + row[24],
    "AV/Equipment Needed: " + row[25],
    "Additional Details: " + row[26]
  ];
  const fullDescription = descriptionLines.join("\n");
  
  Logger.log("Creating event: " + eventName);
  Logger.log("Start: " + Utilities.formatDate(startDate, timeZone, "MM/dd/yyyy hh:mm:ss a"));
  Logger.log("End: " + Utilities.formatDate(endDate, timeZone, "MM/dd/yyyy hh:mm:ss a"));
  
  try {
    calendar.createEvent(eventName, startDate, endDate, { description: fullDescription });
    Logger.log("Event created: " + eventName);
  } catch (error) {
    Logger.log("Error creating event: " + error.message);
  }
}