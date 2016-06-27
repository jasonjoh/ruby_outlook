module RubyOutlook
  class Client

    def get_task_folders(task_group_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/taskgroups/#{task_group_id}" if task_group_id.present? }/taskfolders"
      
      response = make_api_call(:get, request_url)
      JSON.parse(response)
    end

    def get_tasks(task_folder_id = nil, user = nil)
      request_url = "/#{user_or_me(user)}#{"/taskfolders/#{task_folder_id}" if task_folder_id.present? }/tasks"
      
      response = make_api_call(:get, request_url)
      JSON.parse(response)
    end

  end
end
