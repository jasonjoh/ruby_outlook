require "ruby_outlook/version"
require "faraday"
require 'securerandom'
require "json"

module RubyOutlook
  class Client
    # User agent
    attr_reader :user_agent
    # The server to make API calls to.
    # set via .new(api: :graph) etc. See BASE_URLS
    attr_writer :api_host
    attr_writer :enable_fiddler

    # Outlook APIs https://docs.microsoft.com/en-us/outlook/
    #
    # Encouraged API: Graph API v1
    # https://docs.microsoft.com/en-us/graph/use-the-api?view=graph-rest-1.0
    #   https://graph.microsoft.com/{version}
    #   {HTTP method} https://graph.microsoft.com/{version}/{resource}?{query-parameters}
    #   https://graph.microsoft.com/v1.0/$metadata
    #   pagination: response body json property '@odata.nextLink'
    #
    # Previous API: Outlook API v2 (both Outlook.com and Office 365 users)
    # https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api
    #   Most Outlook API features are supported by Graph.
    #   check https://docs.microsoft.com/en-us/outlook/rest/compare-graph for not-yet-supported resources
    #   https://outlook.office.com/api/{version}/{user_context}
    #
    # Previous API: Outlook API v2 (only Office 365 users [Business & School accounts], not Outlook.com [inc Hotmail etc] )
    #   https://outlook.office365.com/api/{version}/{user_context}
    #
    # {version} eg "v1.0" "v2.0" or "beta"
    # {user_context} all of Outlook API and much of Graph API (but not all) is user-scoped.
    #                https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api#target-user
    #                User can be identified like
    #                /me/
    #                /users/{upn}/ eg upn sadie@contoso.com => /users/sadie@contoso.com/
    #                /users/{AAD_userId@AAD_tenandId}/ eg /users/ddfcd489-628b-40d7-b48b-57002df800e5@1717622f-1d94-4d0c-9d74-709fad664b77/
    BASE_URLS = {
        # {HTTP method} https://graph.microsoft.com/{version}/{resource}?{query-parameters}
        # {resource} is prefixed by /me/ or /users/{id or userPrincipalName}/ for user-centric-resources (but not all)
        graph:     'https://graph.microsoft.com/%{version}',
        # {HTTP method} https://outlook.office.com/api/{version}/{user_context}/{resource}?{query_parameters}
        # {HTTP method} https://outlook.office365.com/api/{version}/{user_context}/{resource}?{query_parameters}
        # All outlook resources use {user_context} which is /me/ or /users/{upn userPrincipalName, or AAD_userId@AAD_tenandId}/,
        #   but to interop office/office365 and graph, let's offload {user_context} to {resource}
        office:    'https://outlook.office.com/api/%{version}',
        office365: 'https://outlook.office365.com/api/%{version}',
    }
    API_VERSIONS = {
        graph: %w[v1.0 beta],
        # version support:
        # https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api
        office: %w[v2.0 v1.0 beta],
        office365: %w[v2.0 v1.0 beta],
    }

    # TODO: change the default api endpoint from 'office365' to 'graph', but that could be breaking and deserve a major version update in RubyOutlook
    # api (symbol): One of :graph, :office, :office365. Sets base url if `base_url` omitted and manages
    #               Resource Property Names (:graph uses camelCase, :office/:office365 uses PascalCase)
    # base_url (string): template url for targeted api.
    # version (nil, string): use API version stated, or latest not-beta if nil. see API_VERSIONS
    # return_format (symbol): One of :camel_case, :pascal_case. For forward/backwards compatibility between
    #                         office/office365 (PascalCase) and graph (camelCase) apis.
    #                         NOTE: Graph API accepts both camelCase and PascalCase inputs
    def initialize(api: :office365, base_url: nil, version: nil, return_format: nil, debug: false)
      @user_agent = "RubyOutlookGem/" << RubyOutlook::VERSION

      @api_host = base_url || BASE_URLS[api]
      @version = version || API_VERSIONS[api].first
      # MS Graph API uses cameCase, Outlook REST API used PascalCase
      # ActiveSupport calls these .camelcase(:lower) and .camelcase(:upper) respectively (defaults to :upper ¯\_(ツ)_/¯)
      @resource_format = (api == :graph) ? :camel_case : :pascal_case
      @return_format = return_format || @resource_format

      @enable_fiddler = false
      @debug = debug
    end

    # method (string): The HTTP method to use for the API call.
    #                  Must be 'GET', 'POST', 'PATCH', or 'DELETE'
    # url (string): The URL to use for the API call. Must not contain
    #               the host. For example: '/api/v2.0/me/messages'
    # token (string): access token
    # params (hash) a Ruby hash containing any query parameters needed for the API call
    # payload (hash): a JSON hash representing the API call's payload. Only used
    #                 for POST or PATCH.
    # custom_headers (hash) a Ruby hash of additional headers (eg setting 'x-AnchorMailbox' or 'client-request-id')
    def make_api_call(method, url, token, params = nil, payload = {}, custom_headers=nil)

      conn_params = {
          url: @api_host % { version: @version }
      }

      if @enable_fiddler
        conn_params[:proxy] = 'http://127.0.0.1:8888'
        conn_params[:ssl] = {:verify => false}
      end

      conn = Faraday.new(conn_params) do |faraday|
        # Uses the default Net::HTTP adapter
        faraday.adapter  Faraday.default_adapter
        faraday.response :logger if @debug
      end

      conn.headers = {
        'Authorization' => "Bearer #{token}",
        'Accept' => "application/json",

        # Client instrumentation
        # See https://msdn.microsoft.com/EN-US/library/office/dn720380(v=exchg.150).aspx
        'User-Agent' => @user_agent,
        'client-request-id' => SecureRandom.uuid,
        'return-client-request-id' => "true"
      }

      if custom_headers && custom_headers.class == Hash
        conn.headers = conn.headers.merge( custom_headers )
      end
      
      case method.upcase
        when "GET"
          response = conn.get do |request|
            request.url url, params
          end
        when "POST"
          conn.headers['Content-Type'] = "application/json"
          response = conn.post do |request|
            request.url url, params
            request.body = JSON.dump(payload)
          end
        when "PATCH"
          conn.headers['Content-Type'] = "application/json"
          response = conn.patch do |request|
            request.url url, params
            request.body = JSON.dump(payload)
          end
        when "DELETE"
          response = conn.delete do |request|
            request.url url, params
          end
      end

      if response.status >= 300
        error_info = if response.body.empty?
          ''
          else
            begin
              JSON.parse( response.body )
            rescue JSON::ParserError => _e
              response.body
            end
        end
        return JSON.dump({
          'ruby_outlook_error' => response.status,
          'ruby_outlook_response' => error_info })
      end

      response.body
    end

    #----- Begin Contacts API -----#

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_contacts(token, view_size, page, fields = nil, sort = nil, user = nil)
      request_url = user_context(user) << "/Contacts"
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

      parse_response(get_contacts_response)
    end

    # token (string): access token
    # id (string): The Id of the contact to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_contact_by_id(token, id, fields = nil, user = nil)
      request_url = user_context(user) << "/Contacts/" << id
      request_params = nil

      unless fields.nil?
        request_params = { '$select' => fields.join(',') }
      end

      get_contact_response = make_api_call "GET", request_url, token, request_params

      parse_response(get_contact_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the contact entity
    # folder_id (string): The Id of the contact folder to create the contact in.
    #                     If nil, contact is created in the default contacts folder.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def create_contact(token, payload, folder_id = nil, user = nil)
      request_url = user_context(user)

      unless folder_id.nil?
        request_url << "/ContactFolders/" << folder_id
      end

      request_url << "/Contacts"

      create_contact_response = make_api_call "POST", request_url, token, nil, payload

      parse_response(create_contact_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the updated contact fields
    # id (string): The Id of the contact to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_contact(token, payload, id, user = nil)
      request_url = user_context(user) << "/Contacts/" << id

      update_contact_response = make_api_call "PATCH", request_url, token, nil, payload

      parse_response(update_contact_response)
    end

    # token (string): access token
    # id (string): The Id of the contact to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_contact(token, id, user = nil)
      request_url = user_context(user) << "/Contacts/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

      parse_response(delete_response)
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
      request_url = user_context(user) << "/Messages"
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

      parse_response(get_messages_response)
    end

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    # folder_id (string): The folder to get mail for. (inbox, drafts, sentitems, deleteditems)
    def get_messages_for_folder(token, view_size, page, fields = nil, sort = nil, user = nil, folder_id)
      request_url = user_context(user) << "/MailFolders/#{folder_id}/messages"
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

      parse_response(get_messages_response)
    end

    # token (string): access token
    # id (string): The Id of the message to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_message_by_id(token, id, fields = nil, user = nil)
      request_url = user_context(user) << "/Messages/" << id
      request_params = nil

      unless fields.nil?
        request_params = { '$select' => fields.join(',') }
      end

      get_message_response = make_api_call "GET", request_url, token, request_params

      parse_response(get_message_response)
    end

    # token (string): access token
    # id (string): The Id of the message to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    # returns JSON array of attachments
    def get_attachment_by_message_id(token, id, fields = nil, user = nil)
      request_url = user_context(user) << "/Messages/" << id << "/attachments/"
      request_params = nil

      unless fields.nil?
        request_params = { '$select' => fields.join(',') }
      end

      get_message_response = make_api_call "GET", request_url, token, request_params

      parse_response(get_message_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the contact entity
    # folder_id (string): The Id of the folder to create the message in.
    #                     If nil, message is created in the default drafts folder.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def create_message(token, payload, folder_id = nil, user = nil)
      request_url = user_context(user)

      unless folder_id.nil?
        request_url << "/MailFolders/" << folder_id
      end

      request_url << "/Messages"

      create_message_response = make_api_call "POST", request_url, token, nil, payload

      parse_response(create_message_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the updated message fields
    # id (string): The Id of the message to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_message(token, payload, id, user = nil)
      request_url = user_context(user) << "/Messages/" << id

      update_message_response = make_api_call "PATCH", request_url, token, nil, payload

      parse_response(update_message_response)
    end

    # token (string): access token
    # id (string): The Id of the message to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_message(token, id, user = nil)
      request_url = user_context(user) << "/Messages/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

      parse_response(delete_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the message to send
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def send_message(token, payload, save_to_sentitems = true, user = nil)
      request_url = user_context(user) << "/SendMail"

      # Wrap message in the sendmail JSON structure
      send_mail_json = {
        'Message' => payload,
        'SaveToSentItems' => save_to_sentitems
      }

      send_response = make_api_call "POST", request_url, token, nil, send_mail_json

      return nil if send_response.nil? || send_response.empty?

      parse_response(send_response)
    end

    #----- End Mail API -----#

    #----- Begin Calendar API -----#

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_calendars(token, view_size, page, fields = nil, sort = nil, user = nil)
      request_url = user_context(user) << "/calendars"
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

      parse_response(get_events_response)
    end


    # token (string): access token
    # payload (hash): a JSON hash representing the calendar entity
    #                 {
    #                   "Name": "Social"
    #                 }
    # calendar_group_id (string): The Id of the calendar group to create the calendar in.
    #                     If nil, calendar is created in the default calendar group.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def create_calendar(token, payload, calendar_group_id = nil, user = nil)
      # POST https://outlook.office.com/api/v2.0/me/calendars
      # POST https://outlook.office.com/api/v2.0/me/calendargroups/{calendar_group_id}/calendars

      request_url = user_context(user)

      unless calendar_group_id.nil?
        request_url << "/CalendarGroups/" << calendar_group_id
      end

      request_url << "/calendars"

      create_calendar_response = make_api_call "POST", request_url, token, nil, payload

      parse_response(create_calendar_response)
    end

    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_events(token, view_size, page, fields = nil, sort = nil, user = nil)
      request_url = user_context(user) << "/Events"
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

      parse_response(get_events_response)
    end

    # token (string): access token
    # id (string): The Id of the event to retrieve
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_event_by_id(token, id, fields = nil, user = nil)
      request_url = user_context(user) << "/Events/" << id
      request_params = nil

      unless fields.nil?
        request_params = { '$select' => fields.join(',') }
      end

      get_event_response = make_api_call "GET", request_url, token, request_params

      parse_response(get_event_response)
    end

    # token (string): access token
    # window_start (DateTime): The earliest time (UTC) to include in the view
    # window_end (DateTime): The latest time (UTC) to include in the view
    # id (string): The Id of the calendar to view
    #              If nil, the default calendar is used
    # fields (array): An array of field names to include in results
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_calendar_view(token, window_start, window_end, id = nil, fields = nil, user = nil, limit = 10)
      request_url = user_context(user)

      unless id.nil?
        request_url << "/Calendars/" << id
      end

      request_url << "/CalendarView"

      request_params = {
        'startDateTime' => window_start.strftime('%Y-%m-%dT%H:%M:%SZ'),
        'endDateTime' => window_end.strftime('%Y-%m-%dT%H:%M:%SZ'),
        '$top' => limit
      }

      unless fields.nil?
        request_params['$select'] = fields.join(',')
      end

      get_view_response =make_api_call "GET", request_url, token, request_params

      parse_response(get_view_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the event entity
    # folder_id (string): The Id of the calendar folder to create the event in.
    #                     If nil, event is created in the default calendar folder.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def create_event(token, payload, folder_id = nil, user = nil)
      request_url = user_context(user)

      unless folder_id.nil?
        request_url << "/Calendars/" << folder_id
      end

      request_url << "/Events"

      create_event_response = make_api_call "POST", request_url, token, nil, payload

      parse_response(create_event_response)
    end

    # token (string): access token
    # payload (hash): a JSON hash representing the updated event fields
    # id (string): The Id of the event to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_event(token, payload, id, user = nil)
      request_url = user_context(user) << "/Events/" << id

      update_event_response = make_api_call "PATCH", request_url, token, nil, payload

      parse_response(update_event_response)
    end

    # token (string): access token
    # id (string): The Id of the event to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_event(token, id, user = nil)
      request_url = user_context(user) << "/Events/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

      parse_response(delete_response)
    end

    #----- End Calendar API -----#

    private
    def user_context(user)
      user.nil? ? "Me" : ("users/" << user)
    end

    def parse_response(response)
      # TODO: consider `return nil if response.nil? || response.empty?` for delete_* calls
      parsed = JSON.parse(response)
      parsed = transform_keys(parsed, (@return_format == :camel_case ? :downcase : :upcase)) if @return_format != @resource_format
      parsed
    end

    def transform_keys(response, updown)
      if response.respond_to? :transform_keys!
        response.transform_keys! do |k|
          k[0] ? k.sub(/./) {|c| c.send(updown) } : k
        end
        response.transform_values!{|v| transform_keys(v, updown)}
      end

      response
    end

  end
end
