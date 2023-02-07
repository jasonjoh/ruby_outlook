require "faraday"
require "uuidtools"
require "json"

module RubyOutlook
  class Client
    attr_accessor(*Configuration::VALID_OPTIONS_KEYS)

    def initialize(options = {})
      options = RubyOutlook.options.merge(options)
      Configuration::VALID_OPTIONS_KEYS.each do |key|
        send("#{key}=", options[key])
      end
    end

    def make_api_call(method, path, params = nil, headers = nil, payload = nil)
      handle_response perform_request(method, path, params, headers, payload)
    end

    # private

    def perform_request(method, path, params, headers, payload)
      request_url = endpoint + path

      conn_params = {:url => host}

      if enable_fiddler
        conn_params[:proxy] = 'http://127.0.0.1:8888'
        conn_params[:ssl] = { :verify => false }
      end

      conn = connection(params)
      conn.headers = basic_headers.merge!(headers.to_h)

      case method.to_s.upcase
      when "GET"
        conn.get do |request|
          request.url request_url, params
        end
      when "POST"
        conn.post do |request|
          request.url request_url, params
          request.body = payload.to_json if payload.present?
        end
      when "PATCH"
        conn.patch do |request|
          request.url request_url, params
          request.body = payload.to_json if payload.present?
        end
      when "PUT"
        conn.put do |request|
          request.url request_url, params
          request.body = payload.read if payload.present?
        end
      when "DELETE"
        conn.delete do |request|
          request.url request_url, params
        end
      else
        raise "Unhandled method #{method}"
      end
    end

    def handle_response(response)
      case response.status
      when 200..399
        JSON.parse response.body # TODO: return a return class/object with some helpers
      when 400, 405..409, 411..423
        if response.body.include?('Badly formed token') # this should only happen in a 400
          raise RubyOutlook::SyncStateBadToken.new(response)
        else
          raise RubyOutlook::ClientError.new(response)
        end
      when 410 # official doc says 'gone' - but we only see these for sync tokens that are in a bad state - we'll handle those specifically and throw everything else
        body = response.body
        if body.include?('SyncStateInvalid') || body.include?('SyncStateNotFound')
          raise RubyOutlook::SyncStateInvalid.new(response)
        else
          raise RubyOutlook::ClientError.new(response)
        end
      when 401, 403
        raise RubyOutlook::AuthorizationError.new(response)
      when 404
        raise RubyOutlook::MailboxNotEnabled.new(response) if response.body.include? 'MailboxNotEnabledForRESTAPI'
        raise RubyOutlook::RecordNotFound.new(response)
      when 429
        raise RubyOutlook::RateLimitError.new(response)
      when 500
        raise RubyOutlook::ServerError.new(response)
      when 503
        raise RubyOutlook::ServiceUnavailableError.new(response)
      when 554
        raise RubyOutlook::MailError.new(response)
      else
        raise RubyOutlook::Error.new(response)
      end
    end

    def user_or_me(user)
      user.present? ? "users/#{user}" : "Me"
    end

    def basic_headers
      {
        'Authorization' => "Bearer #{authentication_token}",
        'Accept' => "application/json",

        # Client instrumentation
        # See https://msdn.microsoft.com/EN-US/library/office/dn720380(v=exchg.150).aspx
        'User-Agent' => user_agent,
        'client-request-id' => UUIDTools::UUID.timestamp_create.to_str,
        'return-client-request-id' => "true",
        'Content-Type' => "application/json"
      }
    end

    def connection(params)
      Faraday.new(params) do |faraday|
        # Uses the default Net::HTTP adapter
        faraday.adapter  Faraday.default_adapter
      end
    end

    def build_request_params(params)
      request_params = {}
      request_params['$skiptoken']  = params[:skiptoken]   if params[:skiptoken].present?
      request_params['$deltatoken'] = params[:deltatoken]  if params[:deltatoken].present?
      request_params['$search']     = params[:search]      if params[:search].present?
      request_params['$filter']     = params[:filter]      if params[:filter].present?
      request_params['$select']     = params[:select]      if params[:select].present?
      request_params['$orderby']    = params[:orderby]     if params[:orderby].present?
      request_params['$top']        = params[:top]         if params[:top].present?
      request_params['$skip']       = params[:skip]        if params[:skip].present?
      request_params['$expand']     = params[:expand]      if params[:expand].present?
      request_params['$count']      = params[:count]       if params[:count].present?    # TODO - Check these last couple for correctness
      request_params['$filter']     = "singleValueExtendedProperties/Any(ep:ep/id eq 'String {#{params[:single_value_extended_properties][:guid]}} Name #{params[:single_value_extended_properties][:name]}' and ep/value eq '#{params[:single_value_extended_properties][:value]}')" if params[:single_value_extended_properties].present?

      request_params
    end
  end
end
