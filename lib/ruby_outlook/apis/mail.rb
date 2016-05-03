module RubyOutlook
  class Client
    
    # TODO - fix
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


    def update_message(id, message_attributes, user = nil)
      request_url  = "/#{user_or_me(user)}/Messages/#{id}"

      response = make_api_call(:patch, request_url, nil, nil, message_attributes)
      JSON.parse(response)
    end

    # TODO - fix
    # token (string): access token
    # id (string): The Id of the message to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_message(token, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Messages/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

      JSON.parse(delete_response)
    end

    # TODO - fix
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

  end
end
