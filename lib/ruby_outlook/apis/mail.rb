module RubyOutlook
  class Client

    def get_messages(**args)
      request_params = build_request_params(args)

      make_api_call(:get, "/#{user_or_me(args[:user])}/messages", request_params)
    end

    def get_messages_with_filters(**args)
      request_params = build_request_params(args)
      make_api_call(:get, "/#{user_or_me(args[:user])}/messages", request_params)
    end

    def get_attachments_for(message_id, **args)
      request_url = "/#{user_or_me(args[:user])}/messages/#{message_id}/attachments"
      request_params = build_request_params(args)

      make_api_call(:get, request_url, request_params)
    end

    def get_messages_for_folder(folder_id, **args)
      request_url = "/#{user_or_me(args[:user])}/MailFolders/#{folder_id}/messages"
      request_params = build_request_params(args)

      make_api_call(:get, request_url, request_params)
    end

    # https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0&tabs=http
    def synchonize_messages_for_folder(folder_id, **args)
      request_url = "/#{user_or_me(args[:user])}/MailFolders/#{folder_id}/messages/delta"
      request_params = build_request_params(args)

      headers = {
        'Prefer' => ['odata.track-changes', "odata.maxpagesize=#{args[:max_page_size].presence || 50}"]
      }

      make_api_call(:get, request_url, request_params, headers)
    end

    # https://learn.microsoft.com/en-us/graph/api/mailfolder-delta?view=graph-rest-1.0&tabs=http
    def synchonize_mail_folders(**args)
      request_params = build_request_params(args)

      headers = {
        'Prefer' => ['odata.track-changes', "odata.maxpagesize=#{args[:max_page_size].presence || 50}"]
      }

      make_api_call(:get, "/#{user_or_me(args[:user])}/MailFolders/delta", request_params, headers)
    end

    # https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders?view=graph-rest-1.0&tabs=http
    def get_folders(**args)
      request_params = build_request_params(args)

      headers = {
        'Prefer' => ['odata.track-changes', "odata.maxpagesize=#{args[:max_page_size].presence || 50}"]
      }

      response = make_api_call(:get, "/#{user_or_me(args[:user])}/mailFolders", request_params, headers)
      response if response.present?
    end

    # https://learn.microsoft.com/en-us/graph/api/mailfolder-get?view=graph-rest-1.0&tabs=http
    def get_folder_by_id(folder_id, _fields = nil, user = nil)
      response = make_api_call(:get, "/#{user_or_me(user)}/mailFolders/#{folder_id}")
      response if response.present?
    end

    # https://learn.microsoft.com/en-us/graph/api/mailfolder-list-childfolders?view=graph-rest-1.0&tabs=http
    def get_folder_children_by_id(folder_id, _fields = nil, user = nil)
      make_api_call(:get, "/#{user_or_me(user)}/mailFolders/#{folder_id}/childFolders")
    end

    def get_message_by_id(id, fields = nil, user = nil)
      request_params = fields.present? ? { '$select' => fields.join(',') } : nil

      make_api_call(:get, "/#{user_or_me(user)}/Messages/#{id}", request_params)
    end

    def create_message(message_attributes, folder_id: nil, user: nil)
      request_url = "/#{user_or_me(user)}#{"/folders/#{folder_id}" if folder_id.present? }/messages"

      make_api_call(:post, request_url, nil, nil, message_attributes)
    end

    def create_reply_message(id, comment, message: nil, user: nil)
      reply_json = {
        # TODO - allow for writable message attributes to be set
        "Comment" => comment
      }

      make_api_call(:post, "/#{user_or_me(user)}/messages/#{id}/createreply", nil, nil, reply_json)
    end


    def update_message(id, message_attributes, user = nil)
      make_api_call(:patch, "/#{user_or_me(user)}/Messages/#{id}", nil, nil, message_attributes)
    end

    def delete_message(id, user = nil)
      response = make_api_call(:delete, "/#{user_or_me(user)}/messages/#{id}")
      response if response.present?
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

      send_response
    end

    def send_draft(message_id, user: nil)
      response = make_api_call(:post, "/#{user_or_me(user)}/messages/#{message_id}/send")
      response if response.present?
    end

    # Quick Reply (does not return the response message)
    def reply_all(message_id, comment, message: nil, user: nil)
      request_url = "/#{user_or_me(user)}/messages/#{message_id}/replyall"

      reply_json = {
        # TODO - allow for writable message attributes to be set
        "Comment" => comment
      }

      response = make_api_call(:post, request_url, nil, nil, reply_json)
      response if response.present?
    end

  end
end
