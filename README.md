# RubyOutlook

The RubyOutlook gem is a light-weight implementation of the Office 365 [Mail](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations), [Calendar](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations), and [Contacts](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) REST APIs. It provides basic CRUD functionality for all three APIs, along with the ability to extend functionality by making any arbitrary API call.

For a sample app that uses this gem, see the [Office 365 VCF Import/Export Sample](https://github.com/jasonjoh/o365-vcftool).

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ruby_outlook'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ruby_outlook

## Usage

### Create the client

All functionality is accessed via the `Client` class. Create a new instance of the class to use it:

```ruby
require 'ruby_outlook'
outlook_client = RubyOutlook::Client.new
```

In addition, you can set the `enable_fiddler` property on the `Client` to true if you want to capture Fiddler traces. Setting this property to true sets the proxy for all traffic to `http://127.0.0.1:8888` (the default Fiddler proxy value), and turns off SSL verification. Note that if you set this property to true and do not have Fiddler running, all requests will fail.

### Get an OAuth2 token ###

The Outlook APIs require an OAuth2 token for authentication. This gem doesn't handle the OAuth2 flow for you. For a full example that implements the OAuth2 [Authorization Code Grant Flow](https://msdn.microsoft.com/en-us/library/azure/dn645542.aspx), see the [Office 365 VCF Import/Export Sample](https://github.com/jasonjoh/o365-vcftool).

For convenience, here's the relevant steps and code, which uses the [oauth2](https://rubygems.org/gems/oauth2) gem.

- Generate a login URL:
```ruby
# Generates the login URL for the app.
def get_login_url
  client = OAuth2::Client.new(CLIENT_ID,
                              CLIENT_SECRET,
                              :site => "https://login.microsoftonline.com",
                              :authorize_url => "/common/oauth2/authorize",
                              :token_url => "/common/oauth2/token")

  login_url = client.auth_code.authorize_url(:redirect_uri => "http://yourapp.com/authorize")
end
```
- User browses to the login URL, authenticates, and provides consent to your app.
- The user's browser is redirected back to "http://yourapp.com/authorize", a page in your web app that extracts the `code` parameter from the request URL.
- Exchange the `code` value for an access token:
```ruby
# Exchanges an authorization code for a token
def get_token_from_code(auth_code)
  client = OAuth2::Client.new(CLIENT_ID,
                              CLIENT_SECRET,
                              :site => "https://login.microsoftonline.com",
                              :authorize_url => "/common/oauth2/authorize",
                              :token_url => "/common/oauth2/token")

  token = client.auth_code.get_token(auth_code,
                                     :redirect_uri => "http://yourapp.com/authorize",
                                     :resource => 'https://outlook.office365.com')

  access_token = token
end
```

### Using a built-in function

All of the built-in functions have a required `token` parameter and an optional `user` parameter. The `token` parameter is the OAuth2 access token required for authentication. The `user` parameter is an email address. If passed, the library will make the call to that user's mailbox using the `/Users/user@domain.com/` URL. If omitted, the library will make the call using the `'/Me'` URL.

Other parameters are specific to the call being made. For example, here's how to call `get_contacts`:

```ruby
# A valid access token retrieved via OAuth2
token = 'eyJ0eXAiOiJKV1QiLCJhbGciO...'
# Maximum 30 results per page.
view_size = 30
# Set the page from the query parameter.
page = 1
# Only retrieve display name.
fields = [
"DisplayName"
]
# Sort by display name
sort = { :sort_field => 'DisplayName', :sort_order => 'ASC' }

contacts = outlook_client.get_contacts token,
          view_size, page, fields, sort
```

### Extending functionality

All of the built-in functions wrap the `make_api_call` function. If there is not a built-in function that suits your needs, you can use the `make_api_call` function to implement any API call you want.

```ruby
# method (string): The HTTP method to use for the API call. 
#                  Must be 'GET', 'POST', 'PATCH', or 'DELETE'
# url (string): The URL to use for the API call. Must not contain
#               the host. For example: '/api/v1.0/me/messages'
# token (string): access token
# params (hash) a Ruby hash containing any query parameters needed for the API call
# payload (hash): a JSON hash representing the API call's payload. Only used
#                 for POST or PATCH.
def make_api_call(method, url, token, params = nil, payload = nil)
```

As an example, here's how the library implements `get_contacts`:

```ruby
# token (string): access token
# view_size (int): maximum number of results
# page (int): What page to fetch (multiple of view size)
# fields (array): An array of field names to include in results
# sort (hash): { sort_on => field_to_sort_on, sort_order => 'ASC' | 'DESC' }
# user (string): The user to make the call for. If nil, use the 'Me' constant.
def get_contacts(token, view_size, page, fields = nil, sort = nil, user = nil)
  request_url = "/api/v1.0/" << (user.nil? ? "Me" : ("users/" << user)) << "/Contacts"
  request_params = {
    '$top' => view_size,
    '$skip' => (page - 1) * view_size
  }
  
  if not fields.nil?
    request_params['$select'] = fields.join(',')
  end 
  
  if not sort.nil?
    request_params['$orderby'] = sort[:sort_field] + " " + sort[:sort_order]
  end
  
  get_contacts_response = make_api_call "GET", request_url, token, request_params
  
  return JSON.parse(get_contacts_response)
end
```

Follow the same pattern to implement your own calls.

## Contributing

1. [Fork it](https://github.com/jasonjoh/ruby_outlook/fork)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Exchange Dev Blog](http://blogs.msdn.com/b/exchangedev/)