# multi_cal Overview
multi_cal is a utility built for Calendly and Google Apps Script. Its purpose is to fill in the gaps of Calendly's capabilities and make full use of the Google Suite's services. 

multi_cal enables users to schedule events via Calendly that reserve specific shared resources and which occupy spaces/use resources in a non-blocking manner.

The multi_cal codebase should be readable enough that it explains itself. In this ReadMe, we will go over the steps to setting up multi_cal with a Google account and creating the communication link to your Calendly account.

# Google Apps Script
For multi_cal to work, it needs to be hosted on a dedicated email account. Either select an existing account or create a new one (it will need to have access to the full Google Suite), then navigate to Google Apps Script. Once there, you need to create a new project (you can name it whatever you like) and paste the multi_cal.js code into the editor. Once that is done, you will need to change the three dictionaries (OCCUPANCY_RULES, KEYWORD_TREE, and CALENDAR_TREE) to reflect your setup/purposes.


# Structure Overview
1. All event descriptions in Calendly must include keywords that indicate what shops and/or resources are **in use** during that event. These will follow a specific format.
    - Shop/space and resource/equipment keywords cannot be shared across items that are completely independent of one another (i.e. two identical tools that are in separate shops/spaces will require two distinct keywords).
2. Multiple calendars must always be checked, but **only your personal calendar should be scheduled on.**
    - The scheduler’s calendar is used to check if the individual is available
    - A shop calendar is used to check if the appropriate space or resource is available
3. Each event must be provided free/busy exception rules to make use of keywords.
    - Using exceptions for the keywords described in step 1, we can ensure events in separate spaces do not get falsely flagged as conflicting.
    - For instance, provided there is staff available, an event in KeySpace_1 can always be booked over an event in KeySpace_2. Similarly, if KeySpace_2 has Resource_2 and Resource_3, then Resource_2 time shoul be able to be booked over Resource_1 time.
4. Calendly workflow sends update emails automatically to the shared calendar email. multi_cal hosted on the email account then synchronizes the shared calendar with those emails.
    - multi_cal then adds the keywords from the event description in step 1 to the event name and adds it to the shop calendar.
    - Calendly can then see those keywords in the event name and applies exception rules to the shops calendar. Keywords should NOT be present in your personal events or in titles of events you create through Calendly. This prevents exceptions being applied to personal calendars (which causes double-booking).


## Decide on Your Keywords
Decide keywords for your shops and equipment. Remember, **you’ll need to create an exception for almost every keyword that does not apply to this event.** I recommend making a template event that has all exceptions inside for ease of use. _**Keywords also need to be unique enough that they won’t come up in events you make on your calendar manually.**_ For example: "Space1" is a good keyword, but "Space" is not as it can come up regularly in day-to-day use of Google Calendar.

You can minimize the keywords you need to make exceptions for by using umbrella keywords for whole shops/spaces in addition to keywords for specific resources. For instance, if I am making an event in the KeySpace_1 space, I will only need to add exceptions for the other spaces/umbrella keywords (KeySpace_2, KeySpace_3, KeySpace_4, etc) and for any KeySpace_1 sub-resources/equipment that this event does not conflict with. This method allows me to create a single exception for the umbrella term rather than creating an exception for every resource/equipment that the umbrella contains.


## Initial Setup
In Calendly, you’ll need to enable multiple calendars from your account page. **NOTE: This will apply to all event types you have ever created. Do not take this step if you are not prepared to edit all of your events.**

Go to Availability >> Calendar Settings >> Connect Calendar Account

From there you can connect as many additional calendars as you would like. However, you should still have your **personal calendar selected as the calendar for events to be scheduled on.**

## The Calendar Invite
In Calendly, you will need to make sure that the calendar invitations on your events are set up correctly. 

Ensure that the event has the appropriate **keywords in its description** (more on this later) and that **no keywords are present in the title of the calendar invitation.**

This is done in the Calendar Invitation subsection of the Notification and Workflows section when you edit the event. 

In that subsection, you have the option to change what the event will be called on the calendar and to include specific variables in the event body.

No variables or descriptors you decide to include in the invitation’s title can include keywords or else the system may erroneously double book you. This is why you put keywords in the Event Description instead. When you set up your workflow, these Event Descriptions are what tells multi_cal what keywords to add to the shared-resource events.

![sample_image](images/cal_invite.png)


## Workflows and Google Apps Script
You as a user should not need to use Google Apps Script beyond the initial setup and occasional maintenance of the system. Instead, you will largely employ a set of two workflows in Calendly which will manage the shared calendar(s) for you.

For the system to work appropriately you will need to have the following things done: 
1. All events you create have **keywords in their descriptions** and **no keywords in their calendar name.** Remember that the keywords in the description should describe what resources are being _used_.

    - The last paragraph of the description should be formatted **<< {Duration (in minutes)} | {List of Comma-Separated Keywords} >>** With no new-line breaks. See the example below. You can type any unit signifier for the minutes you like (‘min’, ‘mins’, ‘m’, ‘minutes’, or nothing at all)
      
    - The Exception Rules for the event should be for any keyword that does not conflict with this event. For an event using the KeySpace_1 space, the event should have an exception for every other space offered by your team(s).
    - **Exception rules on round-robin events need to be made for each host.** The creator of the event is not able to make global exceptions.

![sample_image](images/event_description.png)

3. The two workflows you set up must include the relevant event info in an email to the shop calendar:

    - The subject line must start with the action being done. If scheduling a new event, it should start with “new.” If cancelling, it should start with “cancel.” Case does not matter.

    - The workflow must send an email to the host email

    - The workflow must have “send from no-reply” checked

    - The body of the message must follow this format EXACTLY:
    
![sample_image](images/workflows.png)

_**The comma and space between Event Date and Event Time are both required.**_



# The Occupancy Feature (Optional)
multi_cal has occupancy limits hard-coded into a dictionary at the very beginning of the script. It is a constant named OCCUPANCY_RULES.

The exact values can be adjusted and expanded upon, and they are referenced as events are scheduled to make sure no room ever exceeds our desired max occupancy.

To make use of the occupancy feature, simply include bracketed ends directly after a keyword to signify how much of that space or resource they take up. Please note that if you do this for a specific resource, you should also include an occupancy tag for any umbrella keyword it falls under.

Example: KeySpace_1[2] where the 2 represents the attendance/space this event is taking up. There are 3 spots available in the KeySpace_1 space according to the coded dictionary.
An event with this tag should be able to double-book any time where 2 or more spaces remain available in KeySpace_1.

## Calendar Invite Title:
You will need to **include the sharing keyword for the space/resource in the calendar invite title!** Yes, this is an exception to the "Do not include keywords in the title of your calendar invitations" rule.

The convention I use is that share keywords for spaces/umbrellas are the word “Share” + the initials of the space. See below:
  - KeySpace_1 [ShareKS1] - KeySpace_1 Event
  - KeySpace_2 [ShareKS2] - KeySpace_2 Event

For resource sharing, the format is similar, but the convention is “ShareR” + initials:
  - Resource_1 [ShareRR1] - Uses Resource_1 in KeySpace_1
  - Resource_2 [ShareRR2] - Uses Resource_2 in KeySpace_2
  - Resource_3 [ShareRR3] - Uses Resource_3 in KeySpace_2

The code does not look for this convention, so you can adopt any style that you see fit. HOWEVER, **you must be prudent that no share keyword contains another keyword within itself.** Otherwise exceptions may be applied that should not be. 

Generally, the share keyword exists to allow yourself to be double-booked as an individual in addition to allowing spaces/resources to be double-booked (if you do not want this behavior, do not include the share keyword in the calendar invite title). Were it not for the share keyword, a space might allow multiple events to be booked within it but Calendly would look for a different supervisor for each event wanting to use the space. For the event example below, it has two keywords with sharing modifiers.




## Order of Operations: 

**Initial Case:**

A base event gets scheduled. It overlaps no other events. Its title is “Event1 Tag: ShareKS1” and its description keyword is “KeySpace_1[1]”

The script will subtract the event attendance (1, in this example) from the occupancy limit in the OCCUPANCY_RULES dictionary. From there, it will create a string to show many spots are left _available_ in the room by creating the list of attendance sizes that can still be supported: “KeySpace_1[2], KeySpace_1[1]”. This string is then used as the title for a separate event in the calendar to facilitate aligning attendance/occupancy between multiple events. The normal calendar entry for the event will include its original tag (“Event1 Tag: KeySpace_1”) but its free/busy state will be set to “Free” to prevent it from falsely blocking other events.

**Overlaps:**

An overlapping event, “Event 2” with description keyword “KeySpace_1[2]” will be scheduled as “Event2 - KeySpace_1” with its free/busy rule set to “Free”. Any MechEMain occupancy-tracking events that overlap with this new event will be updated with remaining spaces reduced by 2. The script will even split the existing occupancy-tracking events if the overlap is only partial to ensure only the affected time has a reduced availability.

**Exception Rules:**

To allow a space to be double booked: the event booking rules for an occupancy-based keyword require an exception for events with the _same_ keyword occupancy tag.

To allow a single supervisor for multiple events in the same keyword-space, an exception for the relevant sharing keyword is also necessary.

i.e. to get both of the above effects, the event would have the description keyword TAG[X] and would have an exception rule for the same TAG[X] as well as for ShareTAG.

This does not interfere with the normal system, as because TAG[X] includes TAG, anything with a TAG exception will still correctly ignore the TAG[X] events and anything not compatible with ShareTAG will still be blocked.

![sample_image](images/occupancy_exception_rules.png)




# Maintaining the Code
To ensure functionality, there are a couple areas of the code that will need updating whenever the usage expands. They are the three constant dictionaries at the start of the code. First is the OCCUPANCY_RULES dictionary.

This dictionary contains a matched set of keywords with a specific quantifiable occupancy and the integer representing that occupancy limit. If more resources are added requiring occupancy rules or if the occupancy limit of any resources are changed in some other way, this dictionary will need to be edited to reflect that.


Next are the two tree dictionaries.

If any resources are shuffled around or if new calendars are deployed to represent different categories of spaces, these dictionaries will need to be updated. If usage is expanded to include other areas/categories, you will want to create unique keywords for those areas/categories and a separate calendar for them, as well, if necessary.

**Do note that adding a calendar ID to the dictionary does not automatically make this code function.**

_**You will need to share the calendar with the host email address so that it has the permissions required to create events on the calendar.**_
