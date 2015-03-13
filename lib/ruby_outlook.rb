require "ruby_outlook/version"
require "faraday"
require "uuidtools"

module RubyOutlook
  
  class Client
    # User agent
    attr_reader :user_agent
    # The server to make API calls to.
    # Always "https://outlook.office365.com"
    attr_writer :api_host
    attr_writer :enable_fiddler
    
    def make_api_call(method, url, token, params = nil, payload = nil)
      conn_params = {
        :url => 'https://outlook.office365.com'
      }
      
      if @enable_fiddler
        conn_params[:proxy] = 'http://127.0.0.1:8888'
        conn_params[:ssl] = {:verify => false}
      end
    
      conn = Faraday.new(conn_params) do |faraday|
        # Uses the default Net::HTTP adapter
        faraday.adapter  Faraday.default_adapter  
        
      end
      
      conn.headers = {
        'Authorization' => "Bearer #{token}",
        'Accept' => "application/json",
        
        # Client instrumentation
        # See https://msdn.microsoft.com/EN-US/library/office/dn720380(v=exchg.150).aspx
        'User-Agent' => @user_agent,
        'client-request-id' => UUIDTools::UUID.timestamp_create.to_str,
        'return-client-request-id' => "true"
      }
      
      case method.upcase
        when "GET"
          response = conn.get do |request|
            request.url url, params
          end
      end
      
      return response.body
    end
    
    # token (string): access token
    # view_size (int): maximum number of results
    # page (int): What page to fetch (multiple of view size)
    # fields (array): An array of field names to include in results
    # sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
    # user (string): The user to make the call for. If nil, use the 'Me' constant.
    def get_contacts(token, view_size, page, fields, sort = nil, user = nil)
      request_url = "/api/v1.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts"
      request_params = {
        '$top' => view_size,
        '$skip' => (page - 1) * view_size,
        '$select' => fields.join(',')
      }
      
      if not sort.nil?
        request_params['$orderby'] = sort[:sort_field] << " " << sort[:sort_order]
      end
      
      get_contacts_response = make_api_call "GET", request_url, token, request_params
      
      return JSON.parse(get_contacts_response)
    end
    
    private
      def initialize
        @user_agent = "RubyOutlookGem/" << RubyOutlook::VERSION
        @api_host = "https://outlook.office365.com"
        @enable_fiddler = false
        super
      end
  end
end
