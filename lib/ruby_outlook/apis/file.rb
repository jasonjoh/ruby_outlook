module RubyOutlook
  class Client

    def get_drive(**args)
      request_url  = "/#{user_or_me(args[:user])}/drive"
      request_params = build_request_params(args)

      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    def get_items(drive_id, **args)
      request_url = "/#{user_or_me(args[:user])}/drives/#{drive_id}/root/children"
      request_params = build_request_params(args)

      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    def get_children_items(item_id, **args)
      request_url = "/#{user_or_me(args[:user])}/drive/items/#{item_id}/children"
      request_params = build_request_params(args)

      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    def get_item(item_id, **args)
      request_url = "/#{user_or_me(args[:user])}/drive/items/#{item_id}"
      request_params = build_request_params(args)

      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    def delete_item(item_id, user=nil)
      request_url = "/#{user_or_me(user)}/drive/items/#{item_id}"

      response = make_api_call(:delete, request_url)

      return nil if response.blank?
      JSON.parse(response)
    end

    def create_item(path, **args)
      filename = File.basename(path)
      file = File.open(path)
      content_type = MIME::Types.type_for(path).first.content_type
      request_url = "/#{user_or_me(args[:user])}/drive/items/root:/#{filename}:/content"

      response = make_api_call(:put, request_url, { content_type: content_type }, nil, file)
      JSON.parse(response)
    end

    def update_item(item_id, path, **args)
      file = File.open(path)
      content_type = MIME::Types.type_for(path).first.content_type

      request_url = "/#{user_or_me(args[:user])}/drive/items/#{item_id}/content"

      response = make_api_call(:put, request_url, { content_type: content_type }, nil, file)
      JSON.parse(response)
    end

  end
end
