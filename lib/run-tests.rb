# Run from project root like `ruby -Ilib ./lib/run-tests.rb`
require './lib/ruby_outlook.rb'
require 'json'
require 'date' # Needed on Mac to use DateTime

# TODO: Copy a valid, non-expired access token here.
access_token = 'eyJ0eXAiOiJKV1QiLCJhbGciO...'
DEBUG = !ENV['DEBUG'].nil?

# Use CONTACT_ID to GET a pre-existing contact (instead of contact created by test), eg
# 'AAMkADNhMjcxM2U5LWY2MmItNDRjYy05YzgwLWQwY2FmMTU1MjViOABGAAAAAAC_IsPnAGUWR4fYhDeYtiNFBwCDgDrpyW-uTL4a3VuSIF6OAAAAAAEOAACDgDrpyW-uTL4a3VuSIF6OAAAZHKwnAAA='
CONTACT_ID = ENV['CONTACT_ID']

# Use MESSAGE_ID to GET a pre-existing message (instead of message created by test), eg
# 'AAMkADNhMjcxM2U5LWY2MmItNDRjYy05YzgwLWQwY2FmMTU1MjViOABGAAAAAAC_IsPnAGUWR4fYhDeYtiNFBwCDgDrpyW-uTL4a3VuSIF6OAAAAAAEMAACDgDrpyW-uTL4a3VuSIF6OAAAZHKJNAAA='
MESSAGE_ID = ENV['MESSAGE_ID']

# Use EVENT_ID to GET a pre-existing event (instead of event created by test), eg
# 'AAMkADNhMjcxM2U5LWY2MmItNDRjYy05YzgwLWQwY2FmMTU1MjViOABGAAAAAAC_IsPnAGUWR4fYhDeYtiNFBwCDgDrpyW-uTL4a3VuSIF6OAAAAAAENAACDgDrpyW-uTL4a3VuSIF6OAAAXZ15oAAA='
EVENT_ID = ENV['EVENT_ID']


def do_contact_api_tests(token)
  new_contact_payload = <<~JSON % { address: 'pavelb@a830edad9050849NDA1.onmicrosoft.com'}
  {
    "GivenName": "Pavel",
    "Surname": "Bansky",
    "EmailAddresses": [
      {
        "Address": "%{address}",
        "Name": "Pavel Bansky"
      }
    ],
    "BusinessPhones": [
      "+1 732 555 0102"
    ]
  }
  JSON

  update_contact_payload = <<~JSON
  {
    "HomeAddress": {
      "Street": "Some street",
      "City": "Seattle",
      "State": "WA",
      "PostalCode": "98121"
    },
    "Birthday": "1974-07-22"
  }
  JSON

  outlook_client = RubyOutlook::Client.new(return_format: :pascal_case, debug: DEBUG)

  puts 'Testing POST /me/contacts'
  new_contact_json = JSON.parse(new_contact_payload)
  new_contact = outlook_client.create_contact token, new_contact_json
  assert(true, new_contact)

  # Maximum 30 results per page.
  view_size = 30
  # Set the page from the query parameter.
  page = 1
  # Only retrieve display name.
  fields = [
    "DisplayName"
  ]
  # Sort by display name
  sort = { :sort_field => 'DisplayName', :sort_order => 'ASC' }

  puts 'Testing GET /me/contacts'
  contacts = outlook_client.get_contacts token,
              view_size, page, fields, sort
  assert(true, contacts)

  puts 'Testing GET /me/contacts/id'
  contact = outlook_client.get_contact_by_id token, (CONTACT_ID || new_contact['Id'])
  assert(true, contact)

  puts 'Testing PATCH /me/contacts/id'
  update_contact_json = JSON.parse(update_contact_payload)
  updated_contact = outlook_client.update_contact token, update_contact_json, new_contact['Id']
  assert(true, updated_contact)

  puts 'Testing DELETE /me/contacts/id'
  delete_response = outlook_client.delete_contact token, new_contact['Id']
  assert(delete_response.nil?, delete_response)
end

def do_mail_api_tests(token)
  new_message_payload = <<~JSON % { address: "katiej@a830edad9050849NDA1.onmicrosoft.com" }
  {
    "Subject": "Did you see last night's game?",
    "Importance": "Low",
    "Body": {
      "ContentType": "HTML",
      "Content": "They were <b>awesome</b>!"
    },
    "ToRecipients": [
      {
        "EmailAddress": {
          "Address": "%{address}"
        }
      }
    ]
  }
  JSON

  update_message_payload =  <<~JSON
  {
    "Subject": "UPDATED"
  }
  JSON
  
  send_message_payload = <<~JSON % { address: 'allieb@contoso.com'}
  {
    "Subject": "Meet for lunch?",
    "Body": {
      "ContentType": "Text",
      "Content": "The new cafeteria is open."
    },
    "ToRecipients": [
      {
        "EmailAddress": {
          "Address": "%{address}"
        }
      }
    ],
    "Attachments": [
      {
        "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
        "Name": "menu.txt",
        "ContentBytes": "bWFjIGFuZCBjaGVlc2UgdG9kYXk="
      }
    ]
  }
  JSON
  
  outlook_client = RubyOutlook::Client.new(return_format: :pascal_case, debug: DEBUG)

  puts 'Testing POST /me/messages'
  new_message_json = JSON.parse(new_message_payload)
  new_message = outlook_client.create_message token, new_message_json
  assert(true, new_message)

  # Maximum 30 results per page.
  view_size = 30
  # Set the page from the query parameter.
  page = 1
  # Only retrieve display name.
  fields = [
    "Subject"
  ]
  # Sort by display name
  sort = { :sort_field => 'Subject', :sort_order => 'ASC' }

  puts 'Testing GET /me/messages'
  messages = outlook_client.get_messages token,
              view_size, page, fields, sort
  assert(true, messages)

  puts 'Testing GET /me/messages/id'
  message = outlook_client.get_message_by_id token, (MESSAGE_ID || new_message['Id'])
  assert(true, message)

  puts 'Testing PATCH /me/messages/id'
  update_message_json = JSON.parse(update_message_payload)
  updated_message = outlook_client.update_message token, update_message_json, new_message['Id']
  assert(true, updated_message)

  puts 'Testing DELETE /me/messages/id'
  delete_response = outlook_client.delete_message token, new_message['Id']
  assert(delete_response.nil?, delete_response)

  puts 'Testing POST /me/sendmail'
  send_message_json = JSON.parse(send_message_payload)
  send_response = outlook_client.send_message token, send_message_json
  assert(send_response.nil?, send_response)
end

def do_calendar_api_tests(token)
  # NOTE: outlook api v1.0 => v2.0 changed the format of 'Start'/'End' from string-encoded-time to object DateTime/TimeZone
  #   "Start": "2014-07-02T18:00:00Z",
  #   "End": "2014-07-02T19:00:00Z",
  # became
  #       "Start": {
  #           "DateTime": "2014-02-02T18:00:00",
  #           "TimeZone": "Pacific Standard Time"
  #       },
  #       "End": {
  #           "DateTime": "2014-02-02T19:00:00",
  #           "TimeZone": "Pacific Standard Time"
  #       },
  # Having the wrong format returns a http 400 UnableToDeserializePostBody
  new_event_payload = <<~JSON % { address: 'janets@a830edad9050849NDA1.onmicrosoft.com'}
  {
  "Subject": "Discuss the Calendar REST API",
  "Body": {
    "ContentType": "HTML",
    "Content": "I think it will meet our requirements!"
  },
  "Start": {
    "DateTime": "2014-02-02T18:00:00",
    "TimeZone": "Pacific Standard Time"
  },
  "End": {
    "DateTime": "2014-02-02T19:00:00",
    "TimeZone": "Pacific Standard Time"
  },
  "Attendees": [
    {
      "EmailAddress": {
        "Address": "%{address}",
        "Name": "Janet Schorr"
      },
      "Type": "Required"
    }
  ]
  }
  JSON

  update_event_payload = <<~JSON
  {
  "Location": {
    "DisplayName": "Your office"
  }
  }
  JSON
  
  outlook_client = RubyOutlook::Client.new(return_format: :pascal_case, debug: DEBUG)

  # # Tried creating a Calendar to poke at create_event hitting a 400 UnableToDeserializePostBody, unsuccessfully
  # # leaving it commented until gem also supports delete_calendar
  # puts 'Testing POST /me/calendars'
  # new_calendar = outlook_client.create_calendar token, { 'Name' => 'Social'}
  # assert(true, new_calendar)


  puts 'Testing POST /me/events'
  new_event_json = JSON.parse(new_event_payload)
  new_event = outlook_client.create_event token, new_event_json
  assert(true, new_event)

  puts 'Testing GET /me/CalendarView'

  start_time = DateTime.parse('2015-03-03T00:00:00Z')
  end_time = DateTime.parse('2015-03-10T00:00:00Z')
  view = outlook_client.get_calendar_view token, start_time, end_time
  assert(true, view)

    # Maximum 30 results per page.
  view_size = 30
  # Set the page from the query parameter.
  page = 1
  # Only retrieve display name.
  fields = [
    "Subject"
  ]
  # Sort by display name
  sort = { :sort_field => 'Subject', :sort_order => 'ASC' }

  puts 'Testing GET /me/events'
  events = outlook_client.get_events token,
              view_size, page, fields, sort
  assert(true, events)

  puts 'Testing GET /me/events/id'
  event = outlook_client.get_event_by_id token, (EVENT_ID || new_event['Id'])
  assert(true, event)

  puts 'Testing PATCH /me/events/id'
  update_event_json = JSON.parse(update_event_payload)
  updated_event = outlook_client.update_event token, update_event_json, new_event['Id']
  assert(true, updated_event)

  puts 'Testing DELETE /me/events/id'
  delete_response = outlook_client.delete_event token, new_event['Id']
  assert(delete_response.nil?, delete_response)
end


def assert(condition, response = nil)
  puts response
  if condition == false || (response.is_a?(Hash) && !response['ruby_outlook_error'].nil?)
    exit -1
  end

  puts 'SUCCESS'
  puts ''
end

do_contact_api_tests access_token
do_mail_api_tests access_token
do_calendar_api_tests access_token