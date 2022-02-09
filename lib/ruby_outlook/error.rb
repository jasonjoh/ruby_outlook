module RubyOutlook
  # Base error, can wrap another
  class Error < StandardError
    def initialize(response)
      super(formatted_error_message(response))
    end

    private

    def formatted_error_message(response)
      if response.try(:env).present?
        "#{response.env.method.to_s.upcase} #{response.env.url}: #{response.env.status} #{response.env.body}"
      else
        response
      end
    end
  end

  # https://docs.microsoft.com/en-us/graph/errors
  # 400 error
  class ClientError < Error
  end

  # 401 error
  class AuthorizationError < Error
  end

  # 404 error
  class RecordNotFound < Error
  end

  # 429 error - application throttled over mailbox concurrency
  class RateLimitError < Error
  end

  # 500 error
  class ServerError < Error
  end

  # 503 error service unavailable
  class ServiceUnavailableError < Error
  end

  # 554 error - cannot send mail (action needed)
  class MailError < Error
  end
end
