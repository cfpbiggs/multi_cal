// For occupancy-based tags, we need to store the maximum occupancy for the relevant spaces. These numbers may be adjusted as necessary and new spaces may be added as usage expands
const OCCUPANCY_RULES = {
  // These tags are for whole spaces/shops
  "KeySpace_1": 3,
  "KeySpace_2": 6,

  // These tags are for specific resources/equipment. If there is only one of something or only one can be used at a time, it does not need an occupancy entry
  "Resource_1": 2,
  "Resource_2": 6,
  "Resource_3": 3,
};

// Some keywords represent subsets of a larger keyword. For instance Resource_1 is a subset of KeySpace_1. As such, if an event is marked Resource_1 it should also be marked KeySpace_1 so that
// an event looking to book for a separate KeySpace doesn't need to have exceptions for every sub-resource inside KeySpace_1 (or for any other umbrella space not in this example).
const KEYWORD_TREE = {
  // Resource/equipment keywords and their associated space/umbrella keywords
  "Resource_1": "KeySpace_1",
  "Resource_2": "KeySpace_2",
  "Resource_3": "KeySpace_2",
  "Resource_4": "KeySpace_2",
  "Resource_5": "KeySpace_3",
}

// This tree groups the keywords into a physical/departmental category. Each physical/departmental category then has its own calendar ID. Categories do not interface with Calendly, they are simply a means of
// organizing shared calendars in the event there are multiple.
const CALENDAR_TREE = {
  // Keywords and their associated category
  "KeySpace_1": "Category_1",
  "KeySpace_2": "Category_1",
  "KeySpace_3": "Category_1",
  "KeySpace_4": "Category_2",
  "KeySpace_5": "Category_3",
  "KeySpace_6": "Category_3",

  // Categories and their associated calendar ID
  "Category_1" : "<<INSERT CAL_ID FOR CATEGORY_1>>",
  "Category_2" : "<<INSERT CAL_ID FOR CATEGORY_2>>",
  "Category_3" : "<<INSERT CAL_ID FOR CATEGORY_3>>",
}

function readEmail() {
  // This function is recommended to be run every minute. This can be changed in the "Triggers" sidetab of Google Apps Script. 

  // We want to only check for unread Calendly emails, this will save compute time and prevent acknowledging spam.
  var senderEmail = 'no-reply@calendly.com';
  var threads = GmailApp.search('from:' + senderEmail + ' is:unread');  // Search for unread messages from the specific sender
  var shopCalendar = ""; // This where the ID of the calendar is stored once an event needs to be scheduled

  Logger.log("Checking Messages.");

  // Once we have our list of threads to read, we go through all of them
  for (let j = 0; j < threads.length; j++){
    // For each thread, we get a list of messages within that thread (a thread is a list of messages with the same subject line)
    var messages = threads[j].getMessages();

    // Keep our messages in buckets until we are done looping and can clear them away.
    var deleteMessages = [];
    var readMessages = [];
    // Check every message in the thread
    for (let k = 0; k < messages.length; k++){

      // Open a specific message
      var message = messages[k];

      // While the thread may have been marked unread, that doesn't mean that *all* of the messages within the thread were unread. It usually means the most recent message is unread.
      // If this message has already been read, skip it.
      if(!message.isUnread()) continue;

      // Check the subject line of the message
      var subject = message.getSubject().substring(0,3).toUpperCase(); // We only need the first three characters to know what we're doing.
      var body = htmlDeleter(message.getBody()); // Retrieve event information from the HTML block in the body of the message

      // Check that the email body has the right number of lines for an event email. If it doesn't, skip it.
      if (body.length < 3){
        message.markRead();
        Logger.log("A Calendly email was received, but its body was too short.");
        Logger.log(body);
        continue
      }

      // If the email body was long enough, we'll try to actually read out the information it contained.
      else{
        var title, keywords, start, end, des;
        // Try to parse the email body into event details
        try{
          [title, keywords, start, end, des] = parseEmailBody(body);
          
          //Logger.log("Title: " + title);
          //Logger.log("Keywords: " + keywords);
        }
        catch (e) {
          message.markRead();
          Logger.log("Problem interpreting email body. " + e.message);
          Logger.log(body);
          continue
        }

        // Keep track of the adjusted titles used for different events
        var subtitle;

        // We will be nesting another for loop, so we want to be able to detect when we need to skip the message deletion step.
        var deleteFlag = true;

        // To get full use out of Calendly's Free/Busy Exception rules, we need to create/look for a separate calendar event for each keyword on the email event.
        for (let i = 0; i < keywords.length; i++){
          Logger.log("Keyword: " + keywords[i]);
          // Check if this event has an occupancy tag at the end (square brackets containing a number). The function returns the base keyword and -1 for num if there was no occupancy provided
          var [base, num] = separateOccupancy(keywords[i]);
          var busy = true;
          
          // The subtitle for the event should be the title plus what resource is being used.
          subtitle = title + " Using: " + base;

          // Check if the keyword is a subset of a larger keyword. If so, add the superset keyword to the title.
          while(base in KEYWORD_TREE){
            var parent = KEYWORD_TREE[base];
            // add the superset to the title
            subtitle = subtitle + ", " + parent;
            // set base to parent in case the superset is, itself, a subset of another keyword. If it isn't, the loop will break.
            base = parent;
          }

          // Find the relevant shop calendar for the keyword using the new base using the same method as before
          while(base in CALENDAR_TREE){
            var parent = CALENDAR_TREE[base];
            // Set base to parent so that the loop can run again if the parent is the category rather than the calendar ID. If it is the calendar ID, the loop will break.
            base = parent;
          }
          shopCalendar = base;


          // Create an event for this keyword if the email is about scheduling.
          if (subject === "NEW"){
            // If the keyword was occupancy-tagged, check the calendar for other occupancies and make sure the base-only event placed on the calendar is marked as "Free"
            if (num != -1){
              // Adjust the occupancy labels on the calendar for this keyword
              adjustOccupancy(shopCalendar, keywords[i], start, end, false);
              // For occupancy-based events, we set the free/busy of the event to "free"
              busy = false;
            }

            // Try creating a new event using the extracted information. If this does not work, log the body of the email.
            try {
              createEvent(shopCalendar, subtitle, start, end, des, busy);
              Logger.log("Event \"" + subtitle + "\" Created");
            }
            catch (e) {
              Logger.log("Problem creating new event. " + e.message);
              Logger.log(body);
              deleteFlag = false;
            }
          
          }

          // Cancel any existing event for this keyword if the email is about cancelling
          else if (subject === "CAN"){
            // If the keyword was occupancy-tagged, check the calendar for other occupancies and cancel the base-only event.
            if (num != -1){
              adjustOccupancy(shopCalendar, keywords[i], start, end, true);
            }
            // To cancel an event we must first find it on the Google Calendar.
            var eventToModify = findEvent(shopCalendar, subtitle, start, end);
            
            // Try to delete the event (If multiple events or no events are found, the value will be 'null')
            try{
              if (eventToModify == null){
                Logger.log("Event details are valid, but it cannot be cancelled.")
                deleteFlag = false;
              }

              else{
                eventToModify.deleteEvent();
                Logger.log("Event \"" + subtitle + "\" Canceled");
              }
              
            }
            catch(e){
              Logger.log("Problem canceling event. " + e.message);
              Logger.log(body);
              deleteFlag = false;
            }
          }
        }
        
        // Delete the email after logging the output if everything ran correctly
        if (deleteFlag){
          deleteMessages.push(message);
          message.moveToTrash();
          Logger.log("Message Deleted.")
        }
        // If the deleteFlag was turned off, mark message read instead of deleting.
        else{
          readMessages.push(message);
          message.markRead();
          Logger.log("Message Read")
        }
      }
    }
  }

  Logger.log("All Messages Handled.");
}



function parseEmailBody(body){

      // If the user's workflow is set up correctly, the email body will have the following line format:
      // 0 - Event Name
      // 1 - Start Time
      // 2-N - Description
      // N+1 - Extra Event Info: "Duration (in minutes) | Keywords"

      // A RegEx to match the format of << 75mins|KeyWord_1,Resource_1 >> or << 120mins|KeyWord_2,Resource_3 >>
      // Used on the last line of the body

      Logger.log(body);
      var lastIndex = body.length - 1;
      var cleanup = body[lastIndex].split("|");
      Logger.log("Cleanup Array: " + cleanup);
      var unsplitKeys;
      var durationStr = cleanup[0].match(/(\d+).*/i).slice(1);
      Logger.log("Duration Array: " + durationStr);
      var duration = durationStr[0];
      var keywords = [];
      if (cleanup.length > 1){
        // Remove all spaces from the keyword list
        unsplitKeys = cleanup[1].replace(/ /g, "");
        keywords = unsplitKeys.split(",");
      }

      var title = body[0];


      Logger.log("Base Event Title: " + title);

      // The date + start time is in the format Wednesday, June 4, 2025, 1:45pm (Eastern). We need to drop the (Eastern) and the Weekday. 
      // We do this by using split(' ') to get an array with ["Wednesday,", "June", "4,", "2025," , "12:45pm" , "Eastern"]
      // We adjust element 4 (the time) to get rid of the meridian and replace with military time. 
      // We extract elements 1, 2, 3, and 4 to get <June 4, 2025, 13:45>. 
      var time = body[1].split(' ');
      var [hourMinute, meridian] = time[4].match(/(\d{1,2}:\d{2})(am|pm)/i).slice(1); // RegEx will return the original string followed by the parts separated by the expression.
      let [hour, minute] = hourMinute.split(":").map(Number);
      if (meridian.toLowerCase() === "pm" && hour !== 12) hour += 12;
      if (meridian.toLowerCase() === "am" && hour === 12) hour = 0;

      // Get the month's number
      var month = getMonthIndex(time[1]) + 1;
      var day = time[2].slice(0,-1); // cut off the trailing comma
      var year = time[3].slice(0,-1); // cut off the trailing comma


      // Format the date according to the ISO 8601 standard
      var dateString = year + "-" + String(month).padStart(2, '0') + "-" + day.padStart(2, '0') + 'T' + String(hour).padStart(2, '0') + ":"
      + String(minute).padStart(2, "0") + ":00";

      // Create the Date variable (it will default to the script's timezone which can be modified in the settings tab of Apps Script.)
      var start = new Date(dateString);
      // Apply the duration to get the end time.
      var end = new Date(start.getTime());
      end.setMinutes(end.getMinutes() + parseInt(duration));

      // Combine all remaining description lines into one string.
      var des = body.slice(2,lastIndex).join('\n');

      // Return the separated event information
      return [title, keywords, start, end, des];

}



function createEvent(calendarId, title, startTime, endTime, description, busy) {
  // Get the calendar by ID
  var calendar = CalendarApp.getCalendarById(calendarId);

  if (busy){
    // Create the event
    calendar.createEvent(title, new Date(startTime), new Date(endTime), {
      description: description,
  });
  }
  // Create the event, but make sure it is transparent ("Free") so as not to block Calendly needlessly
  else{
    const event = {
      summary: title,
      description: description,
      start: {
        dateTime: startTime.toISOString()
      },
      end: {
        dateTime: endTime.toISOString()
      },
      transparency: "transparent"  //
    };
    // Add the transparent event to the new calendar.
    Calendar.Events.insert(event, calendarId);
  }
  
}



function findEvent(calendarId, title, startTime, endTime){
  // Get the calendar by ID
  var calendar = CalendarApp.getCalendarById(calendarId);

  var start = new Date(startTime.getTime());
  var end = new Date(endTime.getTime());

  // Add a 5 minute buffer to the start and end time
  start.setMinutes(start.getMinutes() - 5);
  end.setMinutes(end.getMinutes() + 5);

  // Find all events in the time range
  var events = calendar.getEvents(start, end);

  // Find any events in the time range that match the title and timing exactly
  toModify = events.filter(event => (event.getTitle().trim() === title.trim() && event.getStartTime().getTime() === startTime.getTime() && event.getEndTime().getTime() === endTime.getTime()));

  // Find how many events match. If it's just one, return the event, if it is more, return nothing.
  var numEvents = toModify.length;
  if (numEvents == 1){
    return toModify[0];
  }
  else if (numEvents == 0){
    Logger.log("No matching events found for title \"" + title + "\"");
    return null;
  }

  else{
    Logger.log("Multiple matching events foundfor title \"" + title + "\"\nReturning the first match.");
    return toModify[0];
  }

}



// For occupancy-notation keywords, we need to be able to separate the occupancy info from the base of the keyword. This function does that and returns the base and a number.
function separateOccupancy(keyword){
  // Find the occupancy usage of the keyword supplied from format {any number of characters} {opening square bracket} {1 or more digits} {closing square bracket}
  var extraction = keyword.match(/^(.*)\[(\d+)\]$/);
  
  // If there was nothing to extract, return the normal keyword and negative 1 (an impossible occupancy number)
  if (!extraction){
    return [keyword,-1];
  }

  // If there WAS something to extract, return the base keyword and the occupancy number given
  return [extraction[1], parseInt(extraction[2], 10)];
}



// CalendarId is a string, keyword is a string, startTime and endTime are Date objects, cancellation is a boolean value {true: this event is being cancelled, false: this event is being created}
function adjustOccupancy(calendarId, keyword, startTime, endTime, cancellation){

  var calendar = CalendarApp.getCalendarById(calendarId);
  
  var [base, num] = separateOccupancy(keyword);
  // Make absolutely certain that there are no leading/trailing spaces
  base = base.trim();
  Logger.log("Keyword Base is: \"" + base + "\"");

  // Check that we actually have an occupancy entry for that base keyword.
  // If not, return the keyword as it was.
  if (!(base in OCCUPANCY_RULES)){
    Logger.log("No occupancy rule was found for keyword \"" + base + "\"");
    return null;
  }

  if (num == -1){
    Logger.log("Keyword did not have occupancy data.");
    return null;
  }

  Logger.log("Checking occupancy from: " + startTime.getTime() + " to: " + endTime.getTime());
  // Get all events for that day
  dayStart = new Date(startTime);
  dayStart.setHours(0,0,0,0);
  dayEnd = new Date(endTime);
  dayEnd.setHours(23,59,59,999);
  
  var events = calendar.getEvents(dayStart, dayEnd);

  // Get all events that overlap with the event being scheduled and which include the keyword in question
  events = events.filter(event => (event.getStartTime() < endTime.getTime() && event.getEndTime() > startTime.getTime() && event.getTitle().includes(base + "[")));
  
  var remaining = OCCUPANCY_RULES[base] - num;

  if (events.length == 0){
    // If this is a cancellation and there are no events to cancel, escape
    if (cancellation){
      return null;
    }
    // If no events are found and we are not cancelling an event, create an event!
    var newEvent = calendar.createEvent(enumerateKeyword(base, remaining), startTime, endTime, {description: remaining})
    if (newEvent.getStartTime().getTime() === newEvent.getEndTime().getTime()){
      newEvent.deleteEvent();
    }
    return null;
  }

  // If this is a cancellation and there are events to adjust, flip the sign of num
  if (cancellation){
    num = -num;
  }
  var early = false;
  var late = false;
  var splitEvents = [];

  // Check against the first conflicting event (we do not loop because we will be doing recursion :) )
  var event = events[0];
  [splitEvents, early, late] = splitEvent(event, startTime, endTime, calendarId);
  // If the existing event contains the new event, we adjust available occupancy in the existing event's middle third.
  if (early && late){
    shiftEvent(splitEvents[1], base, num);
  }
  // If the end of the existing event is overlapped by the start of our new one, adjust occupancy in its later section and run recursion for unchecked second half of new event.
  else if (early){
    adjustOccupancy(calendarId, keyword, splitEvents[1].getEndTime(), endTime, cancellation);
    shiftEvent(splitEvents[1], base, num);
  }
  // If the start of the existing event is overlapped by the end of our new one, adjust occupancy in its earlier section and run recursion for unchecked first half of new event.
  else if (late){
    adjustOccupancy(calendarId, keyword, startTime, splitEvents[0].getStartTime(), cancellation);
    shiftEvent(splitEvents[0], base, num);
  }
  // If our event contained the existing event, reduce occupancy for the existing event and run recursion for the first and last third of the new event.
  else{
    adjustOccupancy(calendarId, keyword, splitEvents[0].getEndTime(), endTime, cancellation);
    adjustOccupancy(calendarId, keyword, startTime, splitEvents[0].getStartTime(), cancellation);
    shiftEvent(splitEvents[0], base, num);
  }
}



function shiftEvent(event, base, num){
  var count = parseInt(event.getDescription(), 10);
  count = count - num;
  // Make sure that no error can make it seem like more space is available than truly exists. If a cancellation would empty the space, shop, or resource, delete the reservation notes for that time.
  if (count >= OCCUPANCY_RULES[base]){
    Logger.log("All reservations removed. Deleting occupancy event.")
    event.deleteEvent();
    return null;
  }
  // Make sure the count never goes negative.
  else if (count < 0){
    count = 0;
  }

  event.setDescription(count);
  event.setTitle(enumerateKeyword(base, count));
}



// Creates a long string enumerating the number of available spaces. This must be done this way to be compatible with Calendly's exception rules.
function enumerateKeyword(base, remaining){

  var concatString = base + "[" + remaining + "]";
  // return adjusted keyword
  for (let i = 1; i < remaining; i++){
    concatString = concatString + ", " + base + "[" + (remaining - i) + "]";
  }

  return concatString
}



// Splits an event using the start and end time given, uses the same name/description for each.
function splitEvent(event, start, end, calendarId){

  var split1 = false;
  var split2 = false;

  var calendar = CalendarApp.getCalendarById(calendarId);

  var events = [event];
  //Logger.log("Event Details:\n" + event.getTitle() + "\n" + event.getStartTime() + "\n" + event.getEndTime() + "\n");

  //Logger.log("Split Details:\n" + start + "\n" + end + "\n");


  if (start > event.getStartTime() && start < event.getEndTime()){
    // If the time DOES cut through the event, tag it 
    split1 = true;
  }

  if (end > event.getStartTime() && end < event.getEndTime()){
    // If the time DOES cut through the event, tag it 
    split2 = true;
  }

  // If the event overlaps both times, we need to split it in three!
  if (split1 && split2){
    // Create two new events! One which spans the middle, one which spans the end
    events.push(calendar.createEvent(event.getTitle(), start, end, {
      description: event.getDescription(),
    }));
    events.push(calendar.createEvent(event.getTitle(), end, event.getEndTime(), {
      description: event.getDescription(),
    }));

    // Adjust the original event to only span the start
    events[0].setTime(event.getStartTime(), time);
  }

  else if (split1){
    // If only the start splits the event, create a second event that spans the latter half.
    events.push(calendar.createEvent(event.getTitle(), start, event.getEndTime(), {
      description: event.getDescription(),
    }));
    events[0].setTime(event.getStartTime(), start);
  }

  else if (split2){
    // Create an event spanning the latter half
    events.push(calendar.createEvent(event.getTitle(), end, event.getEndTime(), {
      description: event.getDescription(),
    }));
    events[0].setTime(event.getStartTime(), end);
  }
  Logger.log([events, split1, split2]);
  // Return the list of events created and a boolean.
  return [events, split1, split2];
}




// This function takes the HTML body of the email and scrapes out the event information from it.
// To troubleshoot this function, print out the output and see if the RegEX is missing any weird clips of HTML or CSS. Add an expression for that missing clip to the end of the replace block.
// If the function is cutting out information you need, check the slice call at the end of the function and check if it needs changing.
function htmlDeleter(html){
  Logger.log(html);
  let newBody = html.replace(/<\/title>/gi, '\n') // Replace </title> with newline
  .replace(/<\/p>/gi, '\n')      // Replace closing </p> with newline
  .replace(/<[^>]*>/g, '')
  .replace(/a\W.*}/g, '')
  .replace(/@media.*{/g, '')
  .replace(/div\W.*}/g, '')
  .replace(/table\Wsocial/g, '')
  .replace(/}/g, '');
  bodyList = newBody.split("\n");
  bodyList = bodyList.map(str => str.trim());
  bodyList = bodyList.filter(str => str !== ""); // Remove all empty strings
  
  
  var bodyLength = bodyList.length;
  var sliceIndex = bodyLength;
  for(let i = 0; i<bodyLength; i++){
    // We go backwards through the list of lines to find the one with the duration and keywords. It is formatted "<Duration> | <Keywords>" so we just need to look for the vertical bar.
    if (bodyList[bodyLength - 1 - i].includes("|")) {
      sliceIndex = bodyLength - i;
      break
    }
  }
  bodyList = bodyList.slice(1,sliceIndex); // Cuts off any trailing lines (Such as "Unsubscribe from Calendly").
  Logger.log(bodyList);
  return bodyList;
}

// This is just a really basic utility function that converts the name of a month into a number from 0-11
function getMonthIndex(monthName) {
  const months = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
  ];
  return months.indexOf(monthName.toLowerCase());
}
