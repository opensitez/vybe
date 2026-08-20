# ============================================================
# REST API Client Framework
# ============================================================
# A full-featured API client with authentication, pagination,
# rate limiting, response caching, middleware pipeline,
# and automatic retry. Models a real-world API wrapper
# like a GitHub, Stripe, or Shopify client.
# ============================================================

require 'net/http'
require 'uri'
require 'json'
require 'base64'
require 'time'
require 'digest/hmac'
require 'openssl'

# ── Middleware pipeline ───────────────────────────────────────────────────────

module Middleware
  class Base
    attr_accessor :next_middleware

    def call(request)
      next_middleware ? next_middleware.call(request) : request
    end
  end

  class Logger < Base
    def call(request)
      start = Time.now
      puts "→ #{request[:method].upcase} #{request[:url]}"
      response = super
      elapsed = ((Time.now - start) * 1000).round(1)
      puts "← #{response[:status]} (#{elapsed}ms)"
      response
    end
  end

  class RateLimiter < Base
    def initialize(requests_per_second: 10)
      @interval = 1.0 / requests_per_second
      @last_call = Time.at(0)
      @mutex = Mutex.new
    end

    def call(request)
      @mutex.synchronize do
        elapsed = Time.now - @last_call
        sleep(@interval - elapsed) if elapsed < @interval
        @last_call = Time.now
      end
      super
    end
  end

  class Retry < Base
    def initialize(max_retries: 3, retry_on: [429, 500, 502, 503, 504])
      @max_retries = max_retries
      @retry_on    = retry_on
    end

    def call(request)
      attempts = 0
      begin
        response = super
        if @retry_on.include?(response[:status]) && attempts < @max_retries
          attempts += 1
          wait = response[:headers]['Retry-After']&.to_i || (2 ** attempts)
          puts "  Retry #{attempts}/#{@max_retries} after #{wait}s (HTTP #{response[:status]})"
          sleep(wait)
          raise "Retrying"
        end
        response
      rescue => e
        retry if e.message == "Retrying" && attempts <= @max_retries
        raise
      end
    end
  end

  class Cache < Base
    def initialize(ttl: 300)
      @store = {}
      @ttl   = ttl
    end

    def call(request)
      return super unless request[:method] == :get

      key = cache_key(request)
      entry = @store[key]

      if entry && Time.now - entry[:cached_at] < @ttl
        puts "  [CACHE HIT] #{request[:url]}"
        return entry[:response]
      end

      response = super
      @store[key] = { response: response, cached_at: Time.now } if response[:status] == 200
      response
    end

    def invalidate(pattern = nil)
      if pattern
        @store.delete_if { |k, _| k.include?(pattern) }
      else
        @store.clear
      end
    end

    private

    def cache_key(request)
      "#{request[:url]}?#{request[:params]&.sort&.to_h}"
    end
  end

  class Authentication < Base
    def initialize(auth_type:, **credentials)
      @auth_type   = auth_type
      @credentials = credentials
    end

    def call(request)
      request[:headers] ||= {}
      case @auth_type
      when :bearer
        request[:headers]['Authorization'] = "Bearer #{@credentials[:token]}"
      when :basic
        encoded = Base64.strict_encode64("#{@credentials[:username]}:#{@credentials[:password]}")
        request[:headers]['Authorization'] = "Basic #{encoded}"
      when :api_key
        if @credentials[:in] == :header
          request[:headers][@credentials[:header] || 'X-API-Key'] = @credentials[:key]
        else
          request[:params] ||= {}
          request[:params][@credentials[:param] || 'api_key'] = @credentials[:key]
        end
      when :hmac
        timestamp = Time.now.to_i.to_s
        payload   = "#{request[:method].upcase}\n#{request[:url]}\n#{timestamp}"
        signature = OpenSSL::HMAC.hexdigest('SHA256', @credentials[:secret], payload)
        request[:headers]['X-Timestamp'] = timestamp
        request[:headers]['X-Signature']  = signature
        request[:headers]['X-API-Key']    = @credentials[:key]
      end
      super
    end
  end
end

# ── HTTP Adapter ──────────────────────────────────────────────────────────────

class HttpAdapter
  def call(request)
    uri = URI(request[:url])
    if request[:params]&.any?
      uri.query = URI.encode_www_form(request[:params])
    end

    http = Net::HTTP.new(uri.host, uri.port)
    http.use_ssl = uri.scheme == 'https'
    http.read_timeout = request[:timeout] || 30

    req = case request[:method]
          when :get    then Net::HTTP::Get.new(uri)
          when :post   then Net::HTTP::Post.new(uri)
          when :put    then Net::HTTP::Put.new(uri)
          when :patch  then Net::HTTP::Patch.new(uri)
          when :delete then Net::HTTP::Delete.new(uri)
          end

    (request[:headers] || {}).each { |k, v| req[k] = v }
    req['Content-Type'] = 'application/json'
    req['Accept']       = 'application/json'

    if request[:body]
      req.body = request[:body].is_a?(String) ? request[:body] : JSON.generate(request[:body])
    end

    begin
      response = http.request(req)
      body = parse_body(response.body)
      {
        status:  response.code.to_i,
        headers: response.to_hash.transform_values(&:first),
        body:    body,
        raw:     response.body
      }
    rescue => e
      { status: 0, headers: {}, body: nil, error: e.message }
    end
  end

  private

  def parse_body(body)
    return nil if body.nil? || body.empty?
    JSON.parse(body, symbolize_names: true)
  rescue JSON::ParserError
    body
  end
end

# ── API Client base ───────────────────────────────────────────────────────────

class ApiClient
  class ApiError < StandardError
    attr_reader :status, :body

    def initialize(message, status: nil, body: nil)
      super(message)
      @status = status
      @body   = body
    end
  end

  class NotFoundError    < ApiError; end
  class UnauthorizedError < ApiError; end
  class RateLimitError   < ApiError; end
  class ServerError      < ApiError; end

  def initialize(base_url:, **options)
    @base_url = base_url.chomp('/')
    @pipeline = build_pipeline(options)
  end

  def get(path, params: nil, **opts)
    request(:get, path, params: params, **opts)
  end

  def post(path, body: nil, **opts)
    request(:post, path, body: body, **opts)
  end

  def put(path, body: nil, **opts)
    request(:put, path, body: body, **opts)
  end

  def patch(path, body: nil, **opts)
    request(:patch, path, body: body, **opts)
  end

  def delete(path, **opts)
    request(:delete, path, **opts)
  end

  # Automatic pagination — yields each page
  def paginate(path, params: {}, page_param: 'page', per_page: 100, &block)
    page = 1
    loop do
      response = get(path, params: params.merge(page_param => page, per_page: per_page))
      items = extract_items(response)
      break if items.empty?
      block.call(items, page)
      break if items.size < per_page
      page += 1
    end
  end

  # Collect all pages into one array
  def get_all(path, params: {}, **opts)
    all_items = []
    paginate(path, params: params, **opts) { |items, _| all_items.concat(items) }
    all_items
  end

  private

  def request(method, path, **opts)
    req = {
      method:  method,
      url:     "#{@base_url}#{path}",
      params:  opts[:params],
      body:    opts[:body],
      headers: opts[:headers] || {}
    }

    response = @pipeline.call(req)
    handle_response(response)
  end

  def handle_response(response)
    case response[:status]
    when 200..299
      response[:body]
    when 401, 403
      raise UnauthorizedError.new("Unauthorized", status: response[:status], body: response[:body])
    when 404
      raise NotFoundError.new("Not Found", status: 404, body: response[:body])
    when 429
      raise RateLimitError.new("Rate Limited", status: 429, body: response[:body])
    when 500..599
      raise ServerError.new("Server Error #{response[:status]}", status: response[:status], body: response[:body])
    else
      raise ApiError.new("HTTP #{response[:status]}", status: response[:status], body: response[:body])
    end
  end

  def extract_items(response)
    case response
    when Array then response
    when Hash  then response[:data] || response[:items] || response[:results] || []
    else []
    end
  end

  def build_pipeline(options)
    middlewares = []
    middlewares << Middleware::Logger.new if options.fetch(:log, true)
    middlewares << Middleware::RateLimiter.new(requests_per_second: options.fetch(:rate_limit, 10))
    middlewares << Middleware::Cache.new(ttl: options.fetch(:cache_ttl, 300)) if options.fetch(:cache, false)
    middlewares << Middleware::Retry.new(max_retries: options.fetch(:max_retries, 3))

    if options[:auth]
      middlewares << Middleware::Authentication.new(**options[:auth])
    end

    adapter = HttpAdapter.new
    middlewares.reverse.reduce(adapter) do |next_mw, mw|
      mw.next_middleware = next_mw
      mw
    end
  end
end

# ── GitHub API client example ─────────────────────────────────────────────────

class GitHubClient < ApiClient
  def initialize(token: nil)
    super(
      base_url:   'https://api.github.com',
      auth:       token ? { auth_type: :bearer, token: token } : nil,
      cache:      true,
      cache_ttl:  60,
      rate_limit: 5,
      log:        true
    )
  end

  def user(username)
    get("/users/#{username}")
  end

  def repos(username, type: 'public', sort: 'updated')
    get_all("/users/#{username}/repos", params: { type: type, sort: sort })
  end

  def repo(owner, name)
    get("/repos/#{owner}/#{name}")
  end

  def issues(owner, repo, state: 'open', labels: nil)
    params = { state: state }
    params[:labels] = labels.join(',') if labels
    get_all("/repos/#{owner}/#{repo}/issues", params: params)
  end

  def create_issue(owner, repo, title:, body: nil, labels: [])
    post("/repos/#{owner}/#{repo}/issues",
         body: { title: title, body: body, labels: labels })
  end

  def rate_limit_status
    get('/rate_limit')
  end
end

# ── Mock API for testing ──────────────────────────────────────────────────────

class MockApiClient < ApiClient
  def initialize
    @responses = {}
    @calls     = []
  end

  def stub(method, path, response:, status: 200)
    @responses["#{method}:#{path}"] = { body: response, status: status }
    self
  end

  def calls_for(path)
    @calls.select { |c| c[:path] == path }
  end

  private

  def request(method, path, **opts)
    @calls << { method: method, path: path, opts: opts }
    key = "#{method}:#{path}"
    stub = @responses[key] || { body: nil, status: 404 }
    raise ApiClient::NotFoundError.new("Not Found") if stub[:status] == 404
    stub[:body]
  end
end

# ── Demo ──────────────────────────────────────────────────────────────────────

if __FILE__ == $0
  puts "=== REST API Client Demo ==="
  puts ""

  # Mock client demo
  mock = MockApiClient.new
  mock.stub(:get, '/users/alice', response: { id: 1, name: 'Alice', email: 'alice@example.com' })
  mock.stub(:get, '/users/alice/repos', response: [
    { name: 'awesome-gem', stars: 142, language: 'Ruby' },
    { name: 'rails-app',   stars: 23,  language: 'Ruby' },
    { name: 'scripts',     stars: 5,   language: 'Shell' }
  ])
  mock.stub(:post, '/users', response: { id: 2, name: 'Bob' }, status: 201)

  user = mock.get('/users/alice')
  puts "User: #{user[:name]} (#{user[:email]})"

  repos = mock.get('/users/alice/repos')
  puts "\nRepositories:"
  repos.each { |r| puts "  #{r[:name]} ★#{r[:stars]} [#{r[:language]}]" }

  puts "\nAPI calls made: #{mock.calls_for('/users/alice').size} to /users/alice"

  # Error handling demo
  puts "\n=== Error Handling ==="
  begin
    mock.get('/nonexistent')
  rescue ApiClient::NotFoundError => e
    puts "Caught NotFoundError: #{e.message}"
  rescue ApiClient::ApiError => e
    puts "Caught ApiError (#{e.status}): #{e.message}"
  end

  # Middleware pipeline demo
  puts "\n=== Middleware Pipeline ==="
  cache = Middleware::Cache.new(ttl: 300)
  rate  = Middleware::RateLimiter.new(requests_per_second: 5)
  puts "Cache middleware: #{cache.class}"
  puts "Rate limiter: #{rate.class}"

  # GitHub client config (no actual requests)
  puts "\n=== GitHub Client Configuration ==="
  gh = GitHubClient.new(token: 'demo_token')
  puts "GitHub client ready: #{gh.class}"
  puts "Base URL: https://api.github.com"
  puts "Features: bearer auth, caching (60s TTL), rate limiting (5 req/s), retry (3x)"
end
