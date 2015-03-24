require './ruby_outlook'
require 'json'

# TODO: Copy a valid, non-expired access token here.
access_token = 'eyJ0eXAiOiJKV1QiLCJhbGciO...'

def do_contact_api_tests(token)
  # TODO: Copy a valid ID for a contact here
  contact_id = 'AAMkADNhMjcxM2U5LWY2MmItNDRjYy05YzgwLWQwY2FmMTU1MjViOABGAAAAAAC_IsPnAGUWR4fYhDeYtiNFBwCDgDrpyW-uTL4a3VuSIF6OAAAAAAEOAACDgDrpyW-uTL4a3VuSIF6OAAAZHKwnAAA='

  new_contact_payload = '{
    "GivenName": "Pavel",
    "Surname": "Bansky",
    "EmailAddresses": [
      {
        "Address": "pavelb@a830edad9050849NDA1.onmicrosoft.com",
        "Name": "Pavel Bansky"
      }
    ],
    "BusinessPhones": [
      "+1 732 555 0102"
    ]
  }'

  update_contact_payload = '{
    "HomeAddress": {
      "Street": "Some street",
      "City": "Seattle",
      "State": "WA",
      "PostalCode": "98121"
    },
    "Birthday": "1974-07-22"
  }'

  outlook_client = RubyOutlook::Client.new

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
              
  puts contacts
  puts ""

  puts 'Testing GET /me/contacts/id'
  contact = outlook_client.get_contact_by_id token, contact_id

  puts contact
  puts ""

  puts 'Testing POST /me/contacts'
  new_contact_json = JSON.parse(new_contact_payload)
  new_contact = outlook_client.create_contact token, new_contact_json

  puts new_contact
  puts ""

  puts 'Testing PATCH /me/contacts/id'
  update_contact_json = JSON.parse(update_contact_payload)
  updated_contact = outlook_client.update_contact token, update_contact_json, new_contact['Id']

  puts updated_contact
  puts ""

  puts 'Testing DELETE /me/contacts/id'
  delete_response = outlook_client.delete_contact token, new_contact['Id']

  puts delete_response.nil? ? "SUCCESS" : delete_response
  puts ""
end

def do_mail_api_tests(token)
  # TODO: Copy a valid ID for a message here
  message_id = 'AAMkADNhMjcxM2U5LWY2MmItNDRjYy05YzgwLWQwY2FmMTU1MjViOABGAAAAAAC_IsPnAGUWR4fYhDeYtiNFBwCDgDrpyW-uTL4a3VuSIF6OAAAAAAEMAACDgDrpyW-uTL4a3VuSIF6OAAAZHKJNAAA='

  new_message_payload = '{
    "Subject": "Did you see last night\'s game?",
    "Importance": "Low",
    "Body": {
      "ContentType": "HTML",
      "Content": "They were <b>awesome</b>!"
    },
    "ToRecipients": [
      {
        "EmailAddress": {
          "Address": "katiej@a830edad9050849NDA1.onmicrosoft.com"
        }
      }
    ]
  }'

  update_message_payload = '{
    "Subject": "UPDATED"
  }'
  
  send_message_payload = '{
    "Subject": "Meet for lunch?",
    "Body": {
      "ContentType": "Text",
      "Content": "The new cafeteria is open."
    },
    "ToRecipients": [
      {
        "EmailAddress": {
          "Address": "allieb@contoso.com"
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
  }'
  
  outlook_client = RubyOutlook::Client.new

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
              
  puts messages
  puts ""

  puts 'Testing GET /me/messages/id'
  message = outlook_client.get_message_by_id token, message_id

  puts message
  puts ""

  puts 'Testing POST /me/messages'
  new_message_json = JSON.parse(new_message_payload)
  new_message = outlook_client.create_message token, new_message_json

  puts new_message
  puts ""

  puts 'Testing PATCH /me/messages/id'
  update_message_json = JSON.parse(update_message_payload)
  updated_message = outlook_client.update_message token, update_message_json, new_message['Id']

  puts updated_message
  puts ""

  puts 'Testing DELETE /me/messages/id'
  delete_response = outlook_client.delete_message token, new_message['Id']

  puts delete_response.nil? ? "SUCCESS" : delete_response
  puts ""
  
  puts 'Testing POST /me/sendmail'
  send_message_json = JSON.parse(send_message_payload)
  send_response = outlook_client.send_message token, send_message_json
  
  puts send_response.nil? ? "SUCCESS" : send_response
  puts ""
end

def do_calendar_api_tests(token)
  # TODO: Copy a valid ID for an event here
  event_id = 'AAMkADNhMjcxM2U5LWY2MmItNDRjYy05YzgwLWQwY2FmMTU1MjViOABGAAAAAAC_IsPnAGUWR4fYhDeYtiNFBwCDgDrpyW-uTL4a3VuSIF6OAAAAAAENAACDgDrpyW-uTL4a3VuSIF6OAAAXZ15oAAA='

  new_event_payload = '{
  "Subject": "Discuss the Calendar REST API",
  "Body": {
    "ContentType": "HTML",
    "Content": "I think it will meet our requirements!"
  },
  "Start": "2014-07-02T18:00:00Z",
  "End": "2014-07-02T19:00:00Z",
  "Attendees": [
    {
      "EmailAddress": {
        "Address": "janets@a830edad9050849NDA1.onmicrosoft.com",
        "Name": "Janet Schorr"
      },
      "Type": "Required"
    }
  ]
}'

  update_event_payload = '{
  "Location": {
    "DisplayName": "Your office"
  }
}'
  
  outlook_client = RubyOutlook::Client.new

  
  puts 'Testing GET /me/CalendarView'
  
  start_time = DateTime.parse('2015-03-03T00:00:00Z')
  end_time = DateTime.parse('2015-03-10T00:00:00Z')
  view = outlook_client.get_calendar_view token, start_time, end_time
              
  puts view
  puts ""
  
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
              
  puts events
  puts ""

  puts 'Testing GET /me/events/id'
  event = outlook_client.get_event_by_id token, event_id

  puts event
  puts ""

  puts 'Testing POST /me/events'
  new_event_json = JSON.parse(new_event_payload)
  new_event = outlook_client.create_event token, new_event_json

  puts new_event
  puts ""

  puts 'Testing PATCH /me/events/id'
  update_event_json = JSON.parse(update_event_payload)
  updated_event = outlook_client.update_event token, update_event_json, new_event['Id']

  puts updated_event
  puts ""

  puts 'Testing DELETE /me/events/id'
  delete_response = outlook_client.delete_event token, new_event['Id']

  puts delete_response.nil? ? "SUCCESS" : delete_response
  puts ""
end 

do_contact_api_tests access_token
do_mail_api_tests access_token
do_calendar_api_tests access_token