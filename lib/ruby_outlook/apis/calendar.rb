module RubyOutlook
  class Client

    def get_calendars(**args)
      request_url  = "/#{user_or_me(args[:user])}#{"/calendargroups/#{args[:calendar_group_id]}" if args[:calendar_group_id].present?}/calendars"
      request_params = build_request_params(args)

      make_api_call(:get, request_url, request_params)
    end

    def sync_events(start_date_time, end_date_time, **args)
      request_url  = "/#{user_or_me(args[:user])}#{"/calendars('#{args[:calendar_id]}')" if args[:calendar_id].present?}/calendarView/delta"
      request_params = build_request_params(args)

      request_params['startDateTime'] = start_date_time.respond_to?(:iso8601) ? start_date_time.iso8601 : start_date_time
      request_params['endDateTime']   = end_date_time.respond_to?(:iso8601)   ? end_date_time.iso8601   : end_date_time

      headers = {
        'Prefer' => ['odata.track-changes', "odata.maxpagesize=#{args[:max_page_size].presence || 50}"]
      }

      make_api_call(:get, request_url, request_params, headers)
    end


    def get_calendar_attachments_for(event_id, **args)
      request_url = "/#{user_or_me(args[:user])}/events/#{event_id}/attachments"
      request_params = build_request_params(args)

      make_api_call(:get, request_url, request_params)
    end

    def create_event(event_attributes, calendar_id: nil, user: nil)
      request_url  = "/#{user_or_me(user)}#{"/calendars('#{calendar_id}')" if calendar_id.present?}/events"

      make_api_call(:post, request_url, nil, nil, event_attributes)
    end

    def update_event(event_id, event_attributes, user: nil)
      make_api_call(:patch, "/#{user_or_me(user)}/events/#{event_id}", nil, nil, event_attributes)
    end

    def delete_event(event_id, user: nil)
      response = make_api_call(:delete, "/#{user_or_me(user)}/events/#{event_id}")
      response if response.present?
    end

    def get_event_by_id(event_id, user: nil)
      response = make_api_call(:get, "/#{user_or_me(user)}/events/#{event_id}")
      response if response.present?
    end

    def respond_to_event(event_id, action, user: nil, comment: nil, send_response: nil)
      action = action.to_s.downcase

      raise RubyOutlook::ClientError.new("#{action} is invalid. Valid actions are accept, tentativelyaccept, or decline.") unless ['accept', 'tentativelyaccept', 'decline'].include?(action)

      update_event_response(action, event_id, user, comment, send_response)
    end

    # TODO - fix
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

      make_api_call "GET", request_url, token, request_params
    end

    # TODO - fix
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

      make_api_call "GET", request_url, token, request_params
    end

    def get_event_instances(start_date_time, end_date_time, **args)
      request_url  = "/#{user_or_me(args[:user])}/events/#{args[:id]}/instances"
      request_params = build_request_params(args)

      request_params['startDateTime'] = start_date_time.respond_to?(:iso8601) ? start_date_time.iso8601 : start_date_time
      request_params['endDateTime']   = end_date_time.respond_to?(:iso8601)   ? end_date_time.iso8601   : end_date_time

      make_api_call(:get, request_url, request_params, request_params)
    end

    private

    def update_event_response(action, event_id, user, comment, send_response)
      request_url = "/#{user_or_me(user)}/events/#{event_id}/#{action}"

      event_attributes = {}
      event_attributes['Comment'] = comment if comment.present?
      event_attributes['SendResponse'] = send_response if [true, false].include?(send_response)

      response = make_api_call(:post, request_url, nil, nil, event_attributes)

      return nil if response.blank?

      response
    end
  end
end
