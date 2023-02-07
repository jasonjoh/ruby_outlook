module RubyOutlook
  class Client

    def joined_teams(**args)
      request_params = build_request_params(args)
      make_api_call(:get, "/#{user_or_me(args[:user])}/joinedTeams", request_params)
    end

    def get_team_members(id)
      make_api_call(:get, "/teams/#{id}/members")
    end

    # https://docs.microsoft.com/en-us/graph/api/team-get-members?view=graph-rest-1.0&tabs=http
    def get_team_member(team_id, member_id)
      make_api_call(:get, "/teams/#{team_id}/members/#{member_id}")
    end

    def get_channels(id)
      make_api_call(:get, "/teams/#{id}/channels")
    end

    def get_channel_messages(team_id, id)
      make_api_call(:get, "/teams/#{team_id}/channels/#{id}/messages")
    end

    def get_message(team_id, channel_id, id)
      make_api_call(:get, "/teams/#{team_id}/channels/#{channel_id}/messages/#{id}")
    end

    def send_reply(team_id, channel_id, id, args)
      request_url = "/teams/#{team_id}/channels/#{channel_id}/messages/#{id}/replies"
      make_api_call(:post, request_url, nil, nil, args)
    end

    def get_replies(team_id, channel_id, id, **args)
      request_params = build_request_params(args)
      make_api_call(:get, "/teams/#{team_id}/channels/#{channel_id}/messages/#{id}/replies", request_params)
    end

    def get_user(id)
      make_api_call(:get, "/users/#{id}")
    end

    def create_channel(team_id, args)
      make_api_call(:post, "/teams/#{team_id}/channels", nil, nil, args)
    end

    def send_message(team_id, channel_id, args)
      make_api_call(:post, "/teams/#{team_id}/channels/#{channel_id}/messages", nil, nil, args)
    end
  end
end
