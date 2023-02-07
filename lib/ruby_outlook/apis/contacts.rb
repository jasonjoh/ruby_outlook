module RubyOutlook
  class Client

    def get_contacts(**args)
      request_params = build_request_params(args)

      make_api_call(:get, "/#{user_or_me(args[:user])}/contacts", request_params)
    end

    def get_contact_by_id(id, select = nil, user = nil)
      request_params = select.present? ? { '$select' => select } : nil

      make_api_call(:get, "/#{user_or_me(user)}/contacts/#{id}", request_params)
    end

    def create_contact(contact_attributes, folder_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/ContactFolders/#{folder_id}" if folder_id.present? }/contacts"

      make_api_call(:post, request_url, nil, nil, contact_attributes)
    end

    def update_contact(id, contact_attributes, user = nil)
      make_api_call(:patch, "/#{user_or_me(user)}/contacts/#{id}", nil, nil, contact_attributes)
    end

    def delete_contact(id, user = nil)
      response = make_api_call(:delete, "/#{user_or_me(user)}/contacts/#{id}")

      return nil if response.blank?

      response
    end

  end
end
