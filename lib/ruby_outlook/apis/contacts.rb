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

    def create_contact(contact_attributes, folder_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/ContactFolders/#{folder_id}" if folder_id.present? }/contacts"

      response = make_api_call(:post, request_url, nil, nil, contact_attributes)
      JSON.parse(response)
    end

    def update_contact(id, contact_attributes, user = nil)
      request_url = "/#{user_or_me(user)}/contacts/#{id}"

      response = make_api_call(:patch, request_url, nil, nil, contact_attributes)
      JSON.parse(response)
    end

    def delete_contact(id, user = nil)
      request_url = "/#{user_or_me(user)}/contacts/#{id}"

      response = make_api_call(:delete, request_url)

      return nil if response.blank?

      JSON.parse(response)
    end
  
  end
end
