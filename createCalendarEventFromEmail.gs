function createCalendarEventFromEmail() {
  // Get unread emails from inbox
  //const threads = GmailApp.search('is:unread');
  var minutes_interval = 10;
  const fiveMinutesAgo = new Date(Date.now() - minutes_interval * 60 * 1000);
  const searchQuery = `is:unread  after:${new Date().toISOString().split('T')[0].replaceAll('-', '/')} before:${new Date(new Date().setDate(new Date().getDate() + 1)).toISOString().split('T')[0].replaceAll('-', '/')}`;
  //console.log(searchQuery);
  const threads = GmailApp.search(searchQuery);

  threads.forEach(thread => {
    const messages = thread.getMessages();

    var email_label;
    labels = thread.getLabels();
    labels.forEach(label => { email_label = label.getName(); });

    if (email_label != 'Calendar Event Created') {


      messages.forEach(message => {
        try {
          // Extract email content

          const subject = message.getSubject();
          //console.log(subject);
          if (subject.toLowerCase().includes('appointment') || subject.toLowerCase().includes('Appointment') || subject.toLowerCase().includes('appt') || subject.toLowerCase().includes('Appt')) {

            const body = message.getPlainBody();
            console.log(subject);
            // Simple date/time extraction regex patterns
            const datePattern = /(?:\d{1,2}\/\d{1,2}\/\d{4}|\d{4}-\d{1,2}-\d{1,2}|\d{4}\/[A-Za-z]{3}\/\d{1,2}|\d{1,2}\/[A-Za-z]{3}\/\d{4}|\d{4}\/[0-9]{1,2}\/\d{1,2}|[A-Za-z]+\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}|(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s+[A-Za-z]{3,5}\s+\d{1,2})/;
            const timePattern = /(?:1[0-2]|0?[1-9])(?::[0-5][0-9])?\s*(?:AM|PM)|(?:[01][0-9]|2[0-3]):[0-5][0-9]/g;

            var dateMatch = body.match(datePattern);
            var timeMatch = body.match(timePattern);

            if (dateMatch == null) {
              dateMatch = subject.match(datePattern);
            }

            if (timeMatch == null) {
              timeMatch = subject.match(timePattern);
            }

            year_str = "" + new Date().getFullYear()
            if (!dateMatch.includes(year_str)) {
              dateMatch = dateMatch + ", " + year_str
            }

            if (typeof(dateMatch) == 'string') {
              dateMatch = [dateMatch]
            }

            //console.log(dateMatch, timeMatch);
            if (dateMatch && timeMatch) {
              // Parse date and time
              const dateStr = dateMatch[0].replace("-", "/");
              const timeStr = timeMatch[0];
              //console.log(dateStr);
              //console.log('converting date');
              // Convert to Date object
              if (timeStr.includes(':')) {
                var [hours, minutes] = timeStr.match(/(\d+):(\d+)/).slice(1);
              } else {
                [hours, minutes] = [parseInt(timeStr), 0];
              }
              const isPM = timeStr.includes('PM');
              const date = new Date(dateStr);
              date.setHours(isPM ? parseInt(hours) + 12 : parseInt(hours));
              date.setMinutes(parseInt(minutes));

              // Round to nearest 15 minutes
              const rounded = new Date(Math.round(date.getTime() / (15 * 60 * 1000)) * (15 * 60 * 1000));

              const eventDate = rounded; //new Date(dateStr + ' ' + timeStr);
              //console.log(eventDate);

              //console.log('adding event to cal');
              // Create calendar event
              
              const event = CalendarApp.getDefaultCalendar().createEvent(
                subject,
                eventDate,
                new Date(eventDate.getTime() + 60 * 60 * 1000), // 1 hour duration
              );
              
              // Set description and guests separately
              event.setDescription(body);
              event.addGuest(message.getFrom());
              
              // Mark email as read
              //message.markRead();
              //console.log('marked read');
              // Add label to processed email
              const label = GmailApp.createLabel('Calendar Event Created');
              thread.addLabel(label);
            }
          }
        } catch (error) {
          console.error('Error processing email:', error);
        }
      });
    }
  });
}

// Create trigger to run script every 5 minutes. Don't need this. Just add a trigger. 
function createTrigger() {
  ScriptApp.newTrigger('createCalendarEventFromEmail')
    .timeBased()
    .everyMinutes(10)
    .create();
}
