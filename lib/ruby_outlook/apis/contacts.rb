module RubyOutlook
  class Client

    def get_contacts(**args)
      request_url = "/#{user_or_me(args[:user])}/contacts"
      request_params = build_request_params(args)
      
      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    def get_contact_by_id(id, select = nil, user = nil)
      request_url = "/#{user_or_me(user)}/contacts/#{id}"

      request_params = select.present? ? { '$select' => select } : nil

      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    # TODO - fix
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

    # TODO - fix
    # token (string): access token
    # payload (hash): a JSON hash representing the updated contact fields
    # id (string): The Id of the contact to update.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def update_contact(token, payload, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts/" << id

      update_contact_response = make_api_call "PATCH", request_url, token, nil, nil, payload

      JSON.parse(update_contact_response)
    end

    # TODO - fix
    # token (string): access token
    # id (string): The Id of the contact to delete.
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def delete_contact(token, id, user = nil)
      request_url = "/api/v2.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts/" << id

      delete_response = make_api_call "DELETE", request_url, token

      return nil if delete_response.nil? || delete_response.empty?

       JSON.parse(delete_response)
    end
  
  end
end
