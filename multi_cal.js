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


// For Group/Enrollment-based events, we keep track of the confirmation window (in minutes), the minimum enrollment, and the maximum enrollment. The key in this dictionary must match the
// title of the corresponding event in Calendly EXACTLY. The confirmation window should also match the setting in Calendly for how far in advance a person must schedule the event.
const ENROLLMENT_RULES = {
  "Group Event 1" : [1440, 2, 3],
  "Group Event 2" : [60, 1, 10],
  "Group Event 3" : [2880, 10, 30],
}


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

  // Group Events and their associated category. Note: Group Events are meant to be public-facing and should not be kept on the same calendar as keyword/occupancy events.
  // Group Events do not include keywords (unless keywords are present in the title), so they may interfere with keyword-based scheduling if put on the same calendar.
  "Group Event 1" : "Category_4",
  "Group Event 2" : "Category_4",
  "Group Event 3" : "Category_5",

  // Categories and their associated calendar ID
  "Category_1" : "<<INSERT CAL_ID FOR CATEGORY_1>>",
  "Category_2" : "<<INSERT CAL_ID FOR CATEGORY_2>>",
  "Category_3" : "<<INSERT CAL_ID FOR CATEGORY_3>>",
  "Category_4" : "<<INSERT CAL_ID FOR CATEGORY_4>>",
  "Category_5" : "<<INSERT CAL_ID FOR CATEGORY_5>>",
}

// These are the three prefixes used on Group/Enrollment based events. They do not need to change.
const GROUP_EVENT_PREFIXES = ["Tentative: ", "Confirmed: ", "Full: "];
// This is the amount of time our automatic group event-checker will look ahead for events in need of cancelling. The longer it is, the longer the checker will take to run. 10080 is 1 week in minutes.
const GROUP_EVENT_LOOKAHEAD = 10080;

function readEmail() {
  // This function is recommended to be run every minute. This can be changed in the "Triggers" sidetab of Google Apps Script. 

  // We want to only check for unread Calendly emails, this will save compute time and prevent acknowledging spam.
  let senderEmail = 'no-reply@calendly.com';
  let threads = GmailApp.search('from:' + senderEmail + ' is:unread');  // Search for unread messages from the specific sender
  let shopCalendar = ""; // This where the ID of the calendar is stored once an event needs to be scheduled

  Logger.log("Checking Messages.");

  // Once we have our list of threads to read, we go through all of them
  for (let j = 0; j < threads.length; j++){
    // For each thread, we get a list of messages within that thread (a thread is a list of messages with the same subject line)
    let messages = threads[j].getMessages();

    // Keep our messages in buckets until we are done looping and can clear them away.
    let deleteMessages = [];
    let readMessages = [];
    // Check every message in the thread
    for (let k = 0; k < messages.length; k++){

      // Open a specific message
      let message = messages[k];

      // While the thread may have been marked unread, that doesn't mean that *all* of the messages within the thread were unread. It usually means the most recent message is unread.
      // If this message has already been read, skip it.
      if(!message.isUnread()) continue;

      // Check the subject line of the message
      let subject = message.getSubject().substring(0,3).toUpperCase(); // We only need the first three characters to know what we're doing.
      let body = htmlDeleter(message.getBody()); // Retrieve event information from the HTML block in the body of the message

      // Check that the email body has the right number of lines for an event email. If it doesn't, skip it.
      if (body.length < 3){
        message.markRead();
        Logger.log("A Calendly email was received, but its body was too short.");
        Logger.log(body);
        continue
      }

      // If the email body was long enough, we'll try to actually read out the information it contained.
      else{
        let title, keywords, start, end, des;
        // Try to parse the email body into event details
        try{
          [title, keywords, start, end, des] = parseEmailBody(body);
        }
        catch (e) {
          message.markRead();
          Logger.log("Problem interpreting email body. " + e.message);
          Logger.log(body);
          continue
        }
        title = title.replace("&amp;", "&");
        // Use the extracted details to add or remove the event from the schedule and record whether the message ought to be deleted.
        let deleteFlag = adjustSchedule(title, subject, keywords, start, end, des);
        
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

// Takes in event details and performs the necessary scheduling actions. Returns a flag to indicate the success of the operation.
function adjustSchedule(title, subject, keywords, start, end, des){
  // Keep track of the adjusted titles used for different events
  let subtitle;
  // This flag keeps track of whether all processes completed successfully.
  let successFlag = true;
  // We will need to keep track of whether this is a Group/Enrollment based event. If it is, the first keyword will be the word "Group." For adjusting occupancies, we'll need to account for
  // only the first person who schedules into this event (if 1 person schedules an event for up to 6 people, we reserve 6 spots until the sign-up time has finished. Once the signup time
  // passes, if enough people signed up, we adjust the occupancy to reflect the number who signed up. If not enough signed up, we fully cancel the 6 reserved spots.)
  // The "Group Event" flag gets set to false only at the top of each message
  let groupFlag = false;
  let groupFlagBlocker = false;
  let groupKeywords;

  if (keywords.length > 0){
    if (keywords[0].toUpperCase() === "GROUP"){
      groupFlag = true;
      groupKeywords = keywords.slice(1,keywords.length);
    }
  }
  
  Logger.log("Group Flag is " + groupFlag);

  // Because we are passing a title by default, no secondary title parameter is needed.
  let groupCal = getCalendar(title, "")[0];

  // If the getCalendar function returned null for the title on a group event, then we have no calendar for this title and should skip it. 
  if (groupCal === null && groupFlag){
    Logger.log("No calendar found for group event \"" + title + "\".")
    return false;
  }
  // This flag is non-blocking (i.e. defaults to True) unless the event is actually a group event and there is an existing entry for it. It must only be reset outside of the keyword loop.
  let newGroupEvent = true;


  // To get full use out of Calendly's Free/Busy Exception rules, we need to create/look for a separate calendar event for each keyword on the email event.
  for (let i = 0; i < keywords.length; i++){
    
    Logger.log("Keyword: " + keywords[i]);
    // Check if this event has an occupancy tag at the end (square brackets containing a number). The function returns the base keyword and -1 for num if there was no occupancy provided
    let [base, num] = separateOccupancy(keywords[i]);
    let busy = true;
    
    // The subtitle for the event should be the title plus what resource is being used.
    subtitle = title + " Using: " + base;

    // Get the calendar for the keyword and check if the keyword is a subset of a larger keyword. If so, add the parent keyword(s) to the title.
    let keyParents = [];
    [shopCalendar, keyParents] = getCalendar(base, title);

    // If there are parent keywords, add them to the subtitle.
    if (keyParents.length > 0){
      subtitle = subtitle + ", " + keyParents.join(", ");
    }
    
    if (shopCalendar === null){
      Logger.log("No calendar found for keyword \"" + base + "\".");
      continue;
    }


    // Create an event for this keyword if the email is about scheduling.
    if (subject === "NEW"){
      let existingGroupEvent;
      if (i == 0 & groupFlag){
        // Check if any event exists for this event item yet but only if we're on the first
        existingGroupEvent = findGroupEvent(groupCal, title, start, end);
        // If there is an existing event, this is not new!
        if (existingGroupEvent != null){
          newGroupEvent = false;
        }
      }
      // If the keyword was occupancy-tagged, check the calendar for other occupancies and make sure the base-only event placed on the calendar is marked as "Free"
      if (num != -1 && newGroupEvent){
        // Adjust the occupancy labels on the calendar for this keyword
        adjustOccupancy(shopCalendar, keywords[i], start, end, false);
        // For occupancy-based events, we set the free/busy of the event to "free"
        busy = false;
      }

      // Try creating a new event using the extracted information. If this does not work, log the body of the email.
      try {
        // If the group flag is enabled and this is, indeed, a new group event, create a Group Event on the calendar.
        if (groupFlag && newGroupEvent && i == 0){
          createGroupEvent(groupCal, title, start, end, des, groupKeywords);
        }
        // If it is a group event but an entry already exists, we must update the old group event (this should only be done if "Group" is the current keyword).
        else if (groupFlag && !newGroupEvent && i == 0){
          Logger.log("Sending Group Event for adjustment:");
          Logger.log(newGroupEvent);
          adjustGroupEvent(existingGroupEvent, true);
        }
        // If it is not a group event or if this is a new group event, create a normal event on the calendar
        else if (!groupFlag || newGroupEvent) {
          createEvent(shopCalendar, subtitle, start, end, des, busy);
          Logger.log("Event \"" + subtitle + "\" Created");
        }
      }
      catch (e) {
        Logger.log("Problem creating new event. " + e.message);
        Logger.log(des);
        successFlag = false;
      }

    }

    // Cancel any existing event for this keyword if the email is about cancelling
    else if (subject === "CAN"){
      // If it's a group event, lower the enrollment as necessary
      if (groupFlag){
        // Check if any event exists for this event item yet.
        let existingGroupEvent = findGroupEvent(groupCal, title, start, end);
        // Adjust the event that was found ONLY if the current keyword is Group (else, we'll knock the event down in enrollment for every keyword it has)
        if (existingGroupEvent != null && i == 0){
          // If the whole group event is cancelled, adjustGroupEvent returns false. If the whole event is cancelled, then we want the full gamut of cancellation processes to run.
          groupFlagBlocker = !adjustGroupEvent(existingGroupEvent, false);
        }
        // If no event was found, log it.
        else if (i == 0) {
          Logger.log("Group Event could not be found for title " + title);
        }
        
      }
      // If it is not a group event, proceed as normal
      else if (!groupFlag) {
        // If the keyword was occupancy-tagged, check the calendar for other occupancies and cancel the base-only event.
        if (num != -1){
          adjustOccupancy(shopCalendar, keywords[i], start, end, true);
        }
        // To cancel an event we must first find it on the Google Calendar.
        let eventToModify = findEvent(shopCalendar, subtitle, start, end);
        
        // Try to delete the event (If multiple events or no events are found, the value will be 'null')
        try{
          if (eventToModify == null){
            Logger.log("Event details are valid, but it cannot be cancelled.")
            successFlag = false;
          }

          else{
            eventToModify.deleteEvent();
            Logger.log("Event \"" + subtitle + "\" Canceled");
          }
          
        }
        catch(e){
          Logger.log("Problem canceling event. " + e.message);
          Logger.log(body);
          successFlag = false;
        }
      }
      // If the whole Group Event was cancelled, turn off the flag for the group event and allow the other keywords to be cancelled as normal.
      if (groupFlagBlocker){
        groupFlag = false;
      }
    }
  }
      return successFlag;
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
      let lastIndex = body.length - 1;
      let cleanup = body[lastIndex].split("|");
      Logger.log("Cleanup Array: " + cleanup);
      let unsplitKeys;
      let durationStr = cleanup[0].match(/(\d+).*/i).slice(1);
      Logger.log("Duration Array: " + durationStr);
      let duration = durationStr[0];
      let keywords = [];
      if (cleanup.length > 1){
        // Remove all spaces from the keyword list
        unsplitKeys = cleanup[1].replace(/ /g, "");
        keywords = unsplitKeys.split(",");
      }

      let title = body[0];


      Logger.log("Base Event Title: " + title);

      // The date + start time is in the format Wednesday, June 4, 2025, 1:45pm (Eastern). We need to drop the (Eastern) and the Weekday. 
      // We do this by using split(' ') to get an array with ["Wednesday,", "June", "4,", "2025," , "12:45pm" , "Eastern"]
      // We adjust element 4 (the time) to get rid of the meridian and replace with military time. 
      // We extract elements 1, 2, 3, and 4 to get <June 4, 2025, 13:45>. 
      let time = body[1].split(' ');
      let [hourMinute, meridian] = time[4].match(/(\d{1,2}:\d{2})(am|pm)/i).slice(1); // RegEx will return the original string followed by the parts separated by the expression.
      let [hour, minute] = hourMinute.split(":").map(Number);
      if (meridian.toLowerCase() === "pm" && hour !== 12) hour += 12;
      if (meridian.toLowerCase() === "am" && hour === 12) hour = 0;

      // Get the month's number
      let month = getMonthIndex(time[1]) + 1;
      let day = time[2].slice(0,-1); // cut off the trailing comma
      let year = time[3].slice(0,-1); // cut off the trailing comma


      // Format the date according to the ISO 8601 standard
      let dateString = year + "-" + String(month).padStart(2, '0') + "-" + day.padStart(2, '0') + 'T' + String(hour).padStart(2, '0') + ":"
      + String(minute).padStart(2, "0") + ":00";

      // Create the Date variable (it will default to the script's timezone which can be modified in the settings tab of Apps Script.)
      let start = new Date(dateString);
      // Apply the duration to get the end time.
      let end = new Date(start.getTime());
      end.setMinutes(end.getMinutes() + parseInt(duration));

      // Combine all remaining description lines into one string.
      let des = body.slice(2,lastIndex).join('\n');

      // Return the separated event information
      return [title, keywords, start, end, des];

}



function createEvent(calendarId, title, startTime, endTime, description, busy) {
  // Get the calendar by ID
  let calendar = CalendarApp.getCalendarById(calendarId);

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


function createGroupEvent(groupEventCalendar, title, startTime, endTime, description, keys){
  Logger.log("Creating Group Event.");
  // Get the calendar by ID
  let calendar = CalendarApp.getCalendarById(groupEventCalendar);
  
  // If there is no enrollment data, do not proceed.
  if(!(title in ENROLLMENT_RULES)){
    Logger.log("Title: " + title + " is not present in the Enrollment Rules Dictionary.");
    return;
  }

  let eventEnrollmentRules = ENROLLMENT_RULES[title];
  let enrollmentDetails = "1 of " + eventEnrollmentRules[2] + " spaces filled.\n";

  let minimum = eventEnrollmentRules[1] - 1;

  // This line was being skipped for some events. Not certain why. It has been reformatted so that it is guaranteed to be run at least in part.
  let minString = minimum + " more sign-ups required to confirm.\n";
  if (minimum === 1){
    minString = "1 more sign-up required to confirm.\n";
  }
  
  enrollmentDetails = enrollmentDetails + minString;

  // Group events are supposed to have a link to the booking page at the end of the description (but before the keywords)
  let descriptionLines = description.split("\n");
  let baseLink = descriptionLines.pop();
  let joinLink = createInstanceLink(baseLink, startTime);
  let signupHook = "Interested in joining? Sign up here:\n" + joinLink + "\n" + title + "\n";
  let newTitle = title;

  if (minimum === 0){
    newTitle = GROUP_EVENT_PREFIXES[1] + newTitle;
  }
  else{
    newTitle = GROUP_EVENT_PREFIXES[0] + newTitle;
  }
  
  // Re-assemble the description now that the link is removed.
  description = descriptionLines.join("\n");

  // Create the event
  calendar.createEvent(newTitle, new Date(startTime), new Date(endTime), {
      description: enrollmentDetails + signupHook + "Tags: " + keys.join(", ") + "\n" + description,
    });
  
}

// Changes the details inside an existing group event. Takes a variable of that existing event and a direction (True = up, False = down) to adjust the enrollment.
// Returns True if the event still exists after adjustment and returns False if the event gets deleted.
function adjustGroupEvent(event, direction){
  Logger.log("Adjusting Group Event details.");
  Logger.log("Adjustment direction is " + direction);
  let description = event.getDescription();
  // Group event description in format
  // 0 Enrollment: # out of [Maximum] spaces filled.
  // 1 Minimum: [Minimum - #] more sign-ups required to confirm..
  // 2 "Interested in joining? Sign up here:"
  // 3 [Signup Link]
  // 4 [Title]
  // 5 Keywords: Tags: [Keywords CSV]
  // 6+ [Description]


  // Extract the details
  let descriptionLines = description.split("\n");

  // If the event is edited by a person, the line separator changes from \n to <br>
  if (descriptionLines.length === 1){
    descriptionLines = descriptionLines[0].split("<br>");
  }
  
  let title = descriptionLines[4];
  let updatedTitle = title;
  let eventEnrollmentRules = ENROLLMENT_RULES[title];
  Logger.log(descriptionLines);
  let keywords = descriptionLines[5].match(/Tags: (.*)/)[1].split(", ");

  // Extract the numbers from lines 0 and 1
  let enrollment = parseInt(descriptionLines[0].match(/(\d*).*/)[1]);

  // The minimum is dependent on the enrollment, so it should not be parsed directly from the text.
  let minimum = eventEnrollmentRules[1] - enrollment;

  // If direction is True, then we are gaining an enrollee
  if (direction){
    enrollment += 1;
    minimum -=1;
  }
  // If direction is False, then we are losing an enrollee. If it's null, then we are updating occupancy for a confirmed event and do not need to adjust enrollment.
  else if (direction !== null) {
    enrollment -= 1;
    minimum += 1;
  }

  let newDescription;

  // if there's no direction to go (i.e. direction is 'null' then we are simply updating a confirmed event whose signup window has passed). No need to do anything fancy.
  if (direction === null){
    // Prevent multi_cal from returning occupancies more than once (i.e. if this function is triggered multiple times within the buffer window)
    if (descriptionLines[descriptionLines.length - 1] === "[Locked]"){
      Logger.log("This event has already been locked.");
      return true;
    }
    Logger.log("Returning unused occupancies for event");
    adjustGroupEventKeywords(title, event.getStartTime(), event.getEndTime(), keywords, enrollment);

    // Now that we've returned any unused occupancies, update the event so that no more people can sign up and "lock" it.
    descriptionLines[2] = "The sign-up window for this event has passed. To book another event, please use the following page:";
    descriptionLines[3] = descriptionLines[3].match(/(.*)\d{4}-\d{2}-\d{2}.*/)[1];
    descriptionLines.push("[Locked]");
    newDescription = descriptionLines.join("\n");
    event.setDescription(newDescription);
    
    return true;
  }
  // If direction was not null, let's update all the event details!
  else{
    descriptionLines[0] = enrollment + " of " + eventEnrollmentRules[2] + " spaces filled.";
    // If there is no minimum number of enrollees left to confirm, change the event title to Confirmed + Title
    if (minimum <= 0){
      descriptionLines[1] = "0 more signups required! The minimum enrollment has been met for this meeting.";
      // Set title to Confirmed
      updatedTitle = GROUP_EVENT_PREFIXES[1] + title;
    }
    // If the minimum number of enrollees hasn't been met yet, keep the title as Tentative and show the number needed for confirmation.
    else if (minimum > 0){
      if (minimum === 1){
        descriptionLines[1] = "1 more sign-up required to confirm.";
      }
      else{
        descriptionLines[1] = minimum + " more sign-ups required to confirm.";
      }
      // set Title to Tentative
      updatedTitle = GROUP_EVENT_PREFIXES[0] + title;
    }
    
    // If the event is full or seems like it might be overbooked, mark it as full and replace the instance-specific signup link with the general booking page.
    if (enrollment >= eventEnrollmentRules[2]){
      descriptionLines[2] = "This event is full! If you need to schedule a training, please book a new training slot using the following link.";
      descriptionLines[3] = descriptionLines[3].match(/(.*)\d{4}-\d{2}-\d{2}.*/)[1];
      updatedTitle = GROUP_EVENT_PREFIXES[2] + title;
    }

    // If the event WAS full but has since been reduced, return the instance-specific signup link and invitation message.
    else if(enrollment == eventEnrollmentRules[2] - 1 && !direction){
      descriptionLines[2] = "Interested in joining? Sign up here:";
      descriptionLines[3] = createInstanceLink(descriptionLines[3], event.getStartTime());
    }

    // If nobody is left enrolled in the event, delete it.
    else if(enrollment == 0){
      adjustGroupEventKeywords(title, event.getStartTime(), event.getEndTime(), keywords, enrollment);
      event.deleteEvent();
      // If the event gets deleted, return False
      return false;
    }
    Logger.log("Changing Group Event details to match enrollment.");
    newDescription = descriptionLines.join("\n");
    Logger.log("New details are:\n" + newDescription);
    event.setTitle(updatedTitle);
    event.setDescription(newDescription);
    // If the event still exists after all that, return True
    return true;
  }
  
}


function adjustGroupEventKeywords(title, start, end, keywords, enrollment){
  // If Enrollment is 0, fully cancel all keyword reservations and return full occupancy reservations.
  // If Enrollment is greater than 0, use ENROLLMENT_RULES to determine how many slots to give back to occupancy-based keywords.
  if (enrollment == 0){
    // If the enrollment drops to 0, go through the full cancellation process for the keywords.
    adjustSchedule(title, "CAN", keywords, start, end, "");
  }
  else{
    let enrollmentDifference = ENROLLMENT_RULES[title][2] - enrollment;
    Logger.log("Maximum occupancy to be returned is " + enrollmentDifference);
    // If enrollment is nonzero, check for occupancy-based keywords and for each of those keywords, add the occupancy back on that has not been signed away.
    for(i = 0; i < keywords.length; i++){
      let [base, num] = separateOccupancy(keywords[i]);
      // If it's an occupancy keyword and the enrollment no longer encompasses everything (some events may have fewer resources than people and they just take turns, for instance), return the appropriate amount of slots.
      if (num != -1 && enrollment < OCCUPANCY_RULES[base]){
        // We're going to "Cancel" the adjusted value which means adding back in the availability for whatever the difference is between the max occupancy (which was reserved) and the actual enrollment
        
        // First find the relevant calendar for the keyword.
        let keywordCal = getCalendar(base, title)[0];
        if (keywordCal === null){
          Logger.log("No calendar found for keyword \"" + base + "\".")
          continue;
        }

        // Next, find the occupancy spaces that can be returned. 
        let occupancyReturn = Math.min(enrollmentDifference, OCCUPANCY_RULES[base]);
        Logger.log("Returning " + occupancyReturn + " to " + base);
        adjustOccupancy(keywordCal, base + "[" + occupancyReturn + "]", start, end, true);
      }
    }
  }
}

// Checks that events meet minimum enrollment ahead of the event according to the provided time-window length. Runs on a trigger every 15 minutes.
function checkEnrollments(){
  // Checks enrollment of group events using the keys from the ENROLLMENT_RULES constant. 
  // Calls the keyword adjuster whenever there is an event whose time is finished.
  Logger.log("Checking for sub-minimum Group Events.");
  titles = Object.keys(ENROLLMENT_RULES);
  
  for (let i = 0; i < titles.length; i++){
    // Find the relevant group calendar for each title. Because we are passing titles directly, we do not need to give a secondary title string.
    let enrolledCalendar = getCalendar(titles[i], "")[0];
    if (enrolledCalendar === null){
      Logger.log("No calendar found for entry \"" + titles[i] + "\".");
      continue;
    }
    // Find the enrollment details for each title
    eventBuffer = ENROLLMENT_RULES[titles[i]][0];

    // Search for tentative events with the same title
    let eventTitle = GROUP_EVENT_PREFIXES[0] + titles[i];
    let events = findUpcomingGroupEvent(enrolledCalendar, eventTitle);
    Logger.log("Found " + events.length + " events to cancel.");
    for (let j = 0; j < events.length; j++){
      // Find the edge of the buffer time for each event
      let bufferTime = events[j].getStartTime();
      bufferTime.setMinutes(bufferTime.getMinutes() - eventBuffer);
      let currentTime = new Date(Date.now());

      // Check the buffer time against the current time. If the current time is later than the edge of the buffer window, the sign-up time has passed, so we can cancel.
      if (bufferTime < currentTime){
        // Create a placeholder event that says "Cancelled: Title" that is set to Free with the same timing.
        createEvent(enrolledCalendar, titles[i], events[j].getStartTime(), events[j].getEndTime(), "", false);

        // Remove the original event
        Logger.log("Event " + events[j].getTitle() + " does not meet its enrollment minimum. It is now being cancelled.");
        let cancelled = false;
        // Reduce the enrollment person by person until the event is cancelled.
        while (!cancelled){
          cancelled = !adjustGroupEvent(events[j], false);
        }
      }
    }

    // Search for confirmed events with the same title
    eventTitle = GROUP_EVENT_PREFIXES[1] + titles[i];
    events = findUpcomingGroupEvent(enrolledCalendar, eventTitle);
    Logger.log("Found " + events.length + " events to update occupancy for.");
    for (let j = 0; j < events.length; j++){
      // Find the edge of the buffer time for each event
      let bufferTime = events[j].getStartTime();
      bufferTime.setMinutes(bufferTime.getMinutes() - eventBuffer);
      let currentTime = new Date(Date.now());
      // Check the buffer time against the current time. If the current time is later than the edge of the buffer window, the sign-up time has passed, so we can cancel.
      
      if (bufferTime < currentTime){
        Logger.log("Event " + events[j].getTitle() + " has met its enrollment minimum but is not full. Any occupancies will be updated accordingly.");
        adjustGroupEvent(events[j], null);
      }
    }
  }
}


// Searches within the GROUP_EVENT_LOOKAHEAD time range for any events of a given title and returns them.
function findUpcomingGroupEvent(calendarID, title){
  // Get the calendar by ID
  let calendar = CalendarApp.getCalendarById(calendarID);
  // Find the start of our search range.
  let startRange = new Date(Date.now());
  // Add a reverse-buffer of 60 minutes in case we are within a cancellation range. This means this function should be triggered MORE OFTEN than once per hour.
  startRange.setMinutes(startRange.getMinutes() - 60);
  let endRange = new Date(startRange.getTime());
  endRange.setMinutes(endRange.getMinutes() + GROUP_EVENT_LOOKAHEAD);
  // Find all events in the time range
  let events = calendar.getEvents(startRange, endRange);

  // Return any events in the time range that match the title
  return events.filter(event => (event.getTitle().trim() === title.trim()));
}

// findEvent gets any events with matching information but only returns the first match--if there are any--and returns null if there are no matches.
function findEvent(calendarId, title, startTime, endTime){
  // Get the calendar by ID
  let calendar = CalendarApp.getCalendarById(calendarId);

  let start = new Date(startTime.getTime());
  let end = new Date(endTime.getTime());

  // Add a 5 minute buffer to the start and end time
  start.setMinutes(start.getMinutes() - 5);
  end.setMinutes(end.getMinutes() + 5);

  // Find all events in the time range
  let events = calendar.getEvents(start, end);

  // Find any events in the time range that match the title and timing exactly
  toModify = events.filter(event => (event.getTitle().trim() === title.trim() && event.getStartTime().getTime() === startTime.getTime() && event.getEndTime().getTime() === endTime.getTime()));

  // Find how many events match. If it's just one, return the event, if it is more, return nothing.
  let numEvents = toModify.length;
  if (numEvents == 1){
    return toModify[0];
  }
  else if (numEvents == 0){
    Logger.log("No matching events found for title \"" + title + "\"");
    return null;
  }

  else{
    Logger.log("Multiple matching events found for title \"" + title + "\"\nReturning the first match.");
    return toModify[0];
  }

}

// findGroupEvent works almost the same as the findEvent function but searches for the three different versions for a particular title that may exist. It prioritizes under-filled events.
function findGroupEvent(calendarID, title, start, end){
  let groupEvent = null;
  for(let i = 0; i < 3; i++){
    // Use the Find Events function to look for an existing entry
    groupEvent = findEvent(calendarID, GROUP_EVENT_PREFIXES[i] + title, start, end);
    // If groupEvents returns "null" it means nothing was found. If groupEvents is NOT null, then there is an event already on the calendar for this item.
    if (groupEvent != null){
      // return the event that was found so it can be used later.
      return groupEvent;
    }
  }
  return groupEvent;
}

// For occupancy-notation keywords, we need to be able to separate the occupancy info from the base of the keyword. This function does that and returns the base and a number.
function separateOccupancy(keyword){
  // Find the occupancy usage of the keyword supplied from format {any number of characters} {opening square bracket} {1 or more digits} {closing square bracket}
  let extraction = keyword.match(/^(.*)\[(\d+)\]$/);
  
  // If there was nothing to extract, return the normal keyword and negative 1 (an impossible occupancy number)
  if (!extraction){
    return [keyword,-1];
  }

  // If there WAS something to extract, return the base keyword and the occupancy number given
  return [extraction[1], parseInt(extraction[2], 10)];
}



// CalendarId is a string, keyword is a string, startTime and endTime are Date objects, cancellation is a boolean value {true: this event is being cancelled, false: this event is being created}
// group Boolean represents whether it is a group event. If it IS, then we only adjust occupancy for it once when it is first made (this reserves the space for the max enrollment of the event) and only when it is fully cancelled or when signup stops does the occupancy get reset or adjusted down to reflect enrollment.
function adjustOccupancy(calendarId, keyword, startTime, endTime, cancellation){

  let calendar = CalendarApp.getCalendarById(calendarId);
  
  let [base, num] = separateOccupancy(keyword);
  // Make absolutely certain that there are no leading/trailing spaces
  base = base.trim();

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

  Logger.log("Checking occupancy for " + base);
  // Get all events for that day
  dayStart = new Date(startTime);
  dayStart.setHours(0,0,0,0);
  dayEnd = new Date(endTime);
  dayEnd.setHours(23,59,59,999);
  
  let events = calendar.getEvents(dayStart, dayEnd);

  // Get all events that overlap with the event being scheduled and which include the keyword in question
  events = events.filter(event => (event.getStartTime() < endTime.getTime() && event.getEndTime() > startTime.getTime() && event.getTitle().includes(base + "[")));
  
  let remaining = OCCUPANCY_RULES[base] - num;

  if (events.length == 0){
    // If this is a cancellation and there are no events to cancel, escape
    if (cancellation){
      return null;
    }
    // If no events are found and we are not cancelling an event, create an event!
    let newEvent = calendar.createEvent(enumerateKeyword(base, remaining), startTime, endTime, {description: remaining})
    if (newEvent.getStartTime().getTime() === newEvent.getEndTime().getTime()){
      newEvent.deleteEvent();
    }
    return null;
  }

  // If this is a cancellation and there are events to adjust, flip the sign of num
  if (cancellation){
    num = -num;
  }
  let early = false;
  let late = false;
  let splitEvents = [];

  // Check against the first conflicting event (we do not loop because we will be doing recursion :) )
  let event = events[0];
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
  let count = parseInt(event.getDescription(), 10);
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

  let concatString = base + "[" + remaining + "]";
  // return adjusted keyword
  for (let i = 1; i < remaining; i++){
    concatString = concatString + ", " + base + "[" + (remaining - i) + "]";
  }

  return concatString
}



// Splits an event using the start and end time given, uses the same name/description for each.
function splitEvent(event, start, end, calendarId){

  let split1 = false;
  let split2 = false;

  let calendar = CalendarApp.getCalendarById(calendarId);

  let events = [event];
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

// This function creates a link to an event instance using the event type's booking page and the time information. The resulting link should look like the following:
// https://calendly.com/user/eventName/2001-01-01T12:00:00
function createInstanceLink(bookingPage, startTime){
  // Use the booking page link as a base
  let instanceLink = bookingPage;
  // If there isn't a slash at the end of the booking page link, add one
  if (instanceLink[instanceLink.length - 1] != "/"){
    instanceLink = instanceLink + "/";
  }
  // Convert the Date/Time into the YYYY-MM-DD and HH:MM:SS formats
  startString = startTime.toISOString();

  return instanceLink + startString.slice(0, startString.length - 5) + "Z";
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
  
  
  let bodyLength = bodyList.length;
  let sliceIndex = bodyLength;
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

// Takes a keyword or event name and finds the calendar ID.
function getCalendar(keyword, title){
  let calID = keyword;
  // if the keyword is "group" then we need to check the calendar tree using the TITLE of the event.
  if(keyword.toUpperCase() === "GROUP"){
    calID = title;
  }
  let superKeys = [];
  while(calID in KEYWORD_TREE){
    superKeys.push(KEYWORD_TREE[calID]);
    calID = KEYWORD_TREE[calID];
  }
  while (calID in CALENDAR_TREE){
    calID = CALENDAR_TREE[calID];
  }

  // Return null if no calendar was found for the entry. All calendar IDs start with "c_"
  if(calID.substring(0,2) != "c_"){
    calID = null;
  }

  return [calID, superKeys];
}
