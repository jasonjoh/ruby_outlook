module RubyOutlook
  module Configuration
    VALID_OPTIONS_KEYS = [
      :authentication_token,
      :enable_fiddler,
      :endpoint,
      :host,
      :user_agent,
    ].freeze

    DEFAULT_AUTHENTICATION_TOKEN = nil
    DEFAULT_ENABLE_FIDDLER       = false
    DEFAULT_ENDPOINT             = '/api/v2.0'.freeze
    DEFAULT_HOST                 = 'https://outlook.office365.com'.freeze
    DEFAULT_USER_AGENT           = "RubyOutlookGem/#{RubyOutlook::VERSION}".freeze

    attr_accessor *VALID_OPTIONS_KEYS

    # Set configuration options to defaults when this module is extended
    def self.extended(base)
      base.reset
    end

    # Allow configuration options to be set in a block
    def configure
      yield self
    end

    def options
      VALID_OPTIONS_KEYS.inject({}) do |option, key|
        option.merge!(key => send(key))
      end
    end

    def reset
      self.authentication_token = DEFAULT_AUTHENTICATION_TOKEN
      self.enable_fiddler       = DEFAULT_ENABLE_FIDDLER
      self.endpoint             = DEFAULT_ENDPOINT
      self.host                 = DEFAULT_HOST
      self.user_agent           = DEFAULT_USER_AGENT
    end
  end
end
