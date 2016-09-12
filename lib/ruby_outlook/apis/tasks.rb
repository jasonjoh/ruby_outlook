module RubyOutlook
  class Client

    def get_task_folders(task_group_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/taskgroups/#{task_group_id}" if task_group_id.present? }/taskfolders"
      
      response = make_api_call(:get, request_url)
      JSON.parse(response)
    end

    def get_tasks(**args)
      request_url = "/#{user_or_me(args[:user])}#{"/taskfolders/#{args[:task_folder_id]}" if args[:task_folder_id].present? }/tasks"
      
      request_params = build_request_params(args)
      
      response = make_api_call(:get, request_url, request_params)
      JSON.parse(response)
    end

    def delete_task(task_id, user = nil)
      request_url = "/#{user_or_me(user)}/tasks/#{task_id}"
      
      response = make_api_call(:delete, request_url)
      JSON.parse(response) if response.present?
    end

    def update_task(task_id, task_attributes, user = nil)
      request_url = "/#{user_or_me(user)}/tasks/#{task_id}"
      
      response = make_api_call(:patch, request_url, nil, nil, task_attributes)
      JSON.parse(response)
    end

    def create_task(task_attributes, task_folder_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/taskfolders/#{task_folder_id}" if task_folder_id.present? }/tasks"
      
      response = make_api_call(:post, request_url, nil, nil, task_attributes)
      JSON.parse(response)
    end

    def create_task_folder(task_folder_attributes, task_group_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/taskgroups/#{task_group_id}" if task_group_id.present? }/taskfolders"

      response = make_api_call(:post, request_url, nil, nil, task_folder_attributes)
      JSON.parse(response)
    end
        
    def update_task_folder(task_folder_id, task_folder_attributes, user = nil)
      request_url = "/#{user_or_me(user)}/taskfolders/#{task_folder_id}"
      
      response = make_api_call(:patch, request_url, nil, nil, task_folder_attributes)
      JSON.parse(response)
    end
    
    def delete_task_folder(task_folder_id, user = nil)
      request_url = "/#{user_or_me(user)}/taskfolders/#{task_folder_id}"
      
      response = make_api_call(:delete, request_url)
      JSON.parse(response) if response.present?
    end
    
  end
end
