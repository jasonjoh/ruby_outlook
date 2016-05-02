require "faraday"
require "uuidtools"
require "json"

# TODO - refactor all methods, move into their own files

module RubyOutlook
  class Client
    attr_accessor(*Configuration::VALID_OPTIONS_KEYS)

    def initialize(options = {})
      options = RubyOutlook.options.merge(options)
      Configuration::VALID_OPTIONS_KEYS.each do |key|
        send("#{key}=", options[key])
      end
    end

    # method (string): The HTTP method to use for the API call.
    #                  Must be 'GET', 'POST', 'PATCH', or 'DELETE'
    # url (string): The URL to use for the API call. Must not contain
    #               the host. For example: '/api/v2.0/me/messages'
    # token (string): access token
    # params (hash) a Ruby hash containing any query parameters needed for the API call
    # payload (hash): a JSON hash representing the API call's payload. Only used
    #                 for POST or PATCH.
    def make_api_call(method, path, params = nil, headers = nil, payload = nil)
      request_url = endpoint + path

      conn_params = {
        :url => host
      }

      if enable_fiddler
        conn_params[:proxy] = 'http://127.0.0.1:8888'
        conn_params[:ssl] = { :verify => false }
      end

      conn = Faraday.new(conn_params) do |faraday|
        # Uses the default Net::HTTP adapter
        faraday.adapter  Faraday.default_adapter
      end

      conn.headers = {
        'Authorization' => "Bearer #{authentication_token}",
        'Accept' => "application/json",
      
        # Client instrumentation
        # See https://msdn.microsoft.com/EN-US/library/office/dn720380(v=exchg.150).aspx
        'User-Agent' => user_agent,
        'client-request-id' => UUIDTools::UUID.timestamp_create.to_str,
        'return-client-request-id' => "true"
      }.merge!(headers.to_h)

      # TODO - symbols
      case method.to_s.upcase
        when "GET"
          response = conn.get do |request|
            request.url request_url, params
          end
        when "POST"
          conn.headers['Content-Type'] = "application/json"
          response = conn.post do |request|
            request.url request_url, params
            request.body = payload.to_json if payload.present?
          end
        when "PATCH"
          conn.headers['Content-Type'] = "application/json"
          response = conn.patch do |request|
            request.url request_url, params
            request.body = payload.to_json if payload.present?
          end
        when "DELETE"
          response = conn.delete do |request|
            request.url request_url, params
          end
      end

      # TODO - remove
      #p response
      p response.headers
      puts ".."
      p response.body
      puts ".."
      #puts response.env
      #puts "===\n\n"

      # To Demo verifcation error
      #
      #raise RubyOutlook::MailError.new("POST https://outlook.office365.com/api/beta/Me/messages/AQMkADAwATNiZmYAZC1lOWE1LTgxZDAtMDACLTAwCgBGAAAD9ZqXqLmCeU_WJdvX85wSmQcA--dVSMxFoUWBXcVxy_0enwAAAgEPAAAA--dVSMxFoUWBXcVxy_0enwAAAAOrd5IAAAA=/send: 554 {\"error\":{\"code\":\"ErrorMessageSubmissionBlocked\",\"message\":\"Cannot send mail. Follow the instructions in your Inbox to verify your account., WASCL UserAction verdict is not None. Actual verdict is HipSend, ShowTierUpgrade.\"}}")

      case response.status
      when 200..399
        response.body
      when 400
        raise RubyOutlook::ClientError.new(response)
      when 401
        raise RubyOutlook::AuthorizationError.new(response)
      when 500
        raise RubyOutlook::ServerError.new(response)
      when 554
        raise RubyOutlook::MailError.new(response)
      else
        raise RubyOutlook::Error.new(response)
      end
    end

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_contacts(token, view_size, page, fields = nil, sort = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts"
      request_params = {
        '$top' => view_size,
        '$skip' => (page - 1) * view_size
      }

      unless fields.nil?
        request_params['$select'] = fields.join(',')
      end

      unless sort.nil?
        request_params['$orderby'] = sort[:sort_field] + " " + sort[:sort_order]
      end

      get_contacts_response = make_api_call "GET", request_url, token, request_params

      JSON.parse(get_contacts_response)
    end

    # token (string): access token
    # id (string): The Id of the contact to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_contact_by_id(token, id, fields = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts/" << id
      request_params = nil

      unless fields.nil?
        request_params = { '$select' => fields.join(',') }
      end

      get_contact_response = make_api_call "GET", request_url, token, request_params

      JSON.parse(get_contact_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the contact entity
    # folder_id (string): The Id of the contact folder to create the contact in.
    #                     If nil, contact is created in the default contacts folder.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def create_contact(token, payload, folder_id = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user))

      unless folder_id.nil?
        request_url << "/ContactFolders/" << folder_id
      end

      request_url << "/Contacts"

      create_contact_response = make_api_call "POST", request_url, token, nil, nil, payload

      JSON.parse(create_contact_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the updated contact fields
    # id (string): The Id of the contact to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_contact(token, payload, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts/" << id

      update_contact_response = make_api_call "PATCH", request_url, token, nil, nil, payload

      JSON.parse(update_contact_response)
    end

    # token (string): access token
    # id (string): The Id of the contact to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_contact(token, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

       JSON.parse(delete_response)
    end

    #----- End Contacts API -----#

    #----- Begin Mail API -----#

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_field: field_to_sort_on, sort_order: 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_messages(token, view_size, page, fields = nil, sort = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Messages"
      request_params = {
        '$top' => view_size,
        '$skip' => (page - 1) * view_size
      }

      unless fields.nil?
        request_params['$select'] = fields.join(',')
      end

      unless sort.nil?
        request_params['$orderby'] = sort[:sort_field] + " " + sort[:sort_order]
      end

      get_messages_response = make_api_call "GET", request_url, token, request_params

      JSON.parse(get_messages_response)
    end

    def get_messages_for_folder(folder_id, **args)
      request_url = "/#{user_or_me(args[:user])}/MailFolders/#{folder_id}/messages"
      request_params = build_request_params(args)
      
      get_messages_response = make_api_call(:get, request_url, request_params)

      JSON.parse(get_messages_response)
    end

    def synchonize_messages_for_folder(folder_id, **args)
      request_url = "/#{user_or_me(args[:user])}/MailFolders/#{folder_id}/messages"
      request_params = build_request_params(args)
      
      headers = {
        'Prefer' => ['odata.track-changes', "odata.maxpagesize=#{args[:max_page_size].presence || 50}"]
      }

      get_messages_response = make_api_call(:get, request_url, request_params, headers)

      JSON.parse(get_messages_response)
    end

    # id (string): The Id of the message to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_message_by_id(id, fields = nil, user = nil)
      request_url  = "/#{user_or_me(user)}/Messages/#{id}"

      request_params = fields.present? ? { '$select' => fields.join(',') } : nil

      get_message_response = make_api_call(:get, request_url, request_params)

      JSON.parse(get_message_response)
    end

    def create_message(message_attributes, folder_id: nil, user: nil)
      request_url = "/#{user_or_me(user)}#{"/folders/#{folder_id}" if folder_id.present? }/messages"
  
      response = make_api_call(:post, request_url, nil, nil, message_attributes)
      JSON.parse(response)
    end

    def create_reply_message(message_id, comment, message: nil, user: nil)
      request_url = "/#{user_or_me(user)}/messages/#{message_id}/createreply"

      reply_json = {
        # TODO - allow for writable message attributes to be set
        "Comment" => comment
      }

      response = make_api_call(:post, request_url, nil, nil, reply_json)
      JSON.parse(response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the updated message fields
    # id (string): The Id of the message to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_message(token, payload, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Messages/" << id

      update_message_response = make_api_call "PATCH", request_url, token, nil, nil, payload

      JSON.parse(update_message_response)
    end

    # token (string): access token
    # id (string): The Id of the message to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_message(token, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Messages/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

      JSON.parse(delete_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the message to send
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def send_message(token, payload, save_to_sentitems = true, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/sendmail"

      # Wrap message in the sendmail JSON structure
      send_mail_json = {
        'Message' => payload,
        'SaveToSentItems' => save_to_sentitems
      }

      send_response = make_api_call "POST", request_url, token, nil, send_mail_json

      return nil if send_response.nil? || send_response.empty?

      JSON.parse(send_response)
    end

    def send_draft(message_id, user: nil)
      request_url = "/#{user_or_me(user)}/messages/#{message_id}/send"
      
      response = make_api_call(:post, request_url)
      JSON.parse(response) if response.present?
    end
  
    # Quick Reply (does not return the response message)
    def reply_all(message_id, comment, message: nil, user: nil)
      request_url = "/#{user_or_me(user)}/messages/#{message_id}/replyall"

      reply_json = {
        # TODO - allow for writable message attributes to be set
        "Comment" => comment
      }

      response = make_api_call(:post, request_url, nil, nil, reply_json)
      JSON.parse(response) if response.present?
    end

    #----- End Mail API -----#

    #----- Begin Calendar API -----#

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_events(token, view_size, page, fields = nil, sort = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Events"
      request_params = {
        '$top' => view_size,
        '$skip' => (page - 1) * view_size
      }

      unless fields.nil?
        request_params['$select'] = fields.join(',')
      end

      unless sort.nil?
        request_params['$orderby'] = sort[:sort_field] + " " + sort[:sort_order]
      end

      get_events_response = make_api_call "GET", request_url, token, request_params

      JSON.parse(get_events_response)
    end

    # token (string): access token
    # id (string): The Id of the event to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_event_by_id(token, id, fields = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Events/" << id
      request_params = nil

      unless fields.nil?
        request_params = { '$select' => fields.join(',') }
      end

      get_event_response = make_api_call "GET", request_url, token, request_params

      JSON.parse(get_event_response)
    end

    # token (string): access token
    # window_start (DateTime): The earliest time (UTC) to include in the view
    # window_end (DateTime): The latest time (UTC) to include in the view
    # id (string): The Id of the calendar to view
    #              If nil, the default calendar is used
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_calendar_view(token, window_start, window_end, id = nil, fields = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user))

      unless id.nil?
        request_url << "/Calendars/" << id
      end

      request_url << "/CalendarView"

      request_params = {
        'startDateTime' => window_start.strftime('%Y-%m-%dT00:00:00Z'),
        'endDateTime' => window_end.strftime('%Y-%m-%dT00:00:00Z')
      }

      unless fields.nil?
        request_params['$select'] = fields.join(',')
      end

      get_view_response =make_api_call "GET", request_url, token, request_params

      JSON.parse(get_view_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the event entity
    # folder_id (string): The Id of the calendar folder to create the event in.
    #                     If nil, event is created in the default calendar folder.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def create_event(token, payload, folder_id = nil, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user))

      unless folder_id.nil?
        request_url << "/Calendars/" << folder_id
      end

      request_url << "/Events"

      create_event_response = make_api_call "POST", request_url, token, nil, nil, payload

      JSON.parse(create_event_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the updated event fields
    # id (string): The Id of the event to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_event(token, payload, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Events/" << id

      update_event_response = make_api_call "PATCH", request_url, token, nil, nil, payload

      JSON.parse(update_event_response)
    end

    # token (string): access token
    # id (string): The Id of the event to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_event(token, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Events/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

      JSON.parse(delete_response)
    end

    private

    def user_or_me(user)
      user.present? ? "users/#{user}" : "Me"
    end

    def build_request_params(params)
      request_params = {}
      request_params['$skiptoken']  = params[:skiptoken]   if params[:skiptoken].present?
      request_params['$deltatoken'] = params[:deltatoken]  if params[:deltatoken].present?
      request_params['$search']     = params[:search]      if params[:search].present?
      request_params['$filter']     = params[:filter]      if params[:filter].present?
      request_params['$select']     = params[:select]      if params[:select].present?
      request_params['$orderby']    = params[:orderby]     if params[:orderby].present?
      request_params['$top']        = params[:top]         if params[:top].present?
      request_params['$skip']       = params[:skip]        if params[:skip].present?
      request_params['$expand']     = params[:expand]      if params[:expand].present?
      request_params['$count']      = params[:count]       if params[:count].present?    # TODO - Check these last couple for correctness
      
      request_params
    end
  end
end
