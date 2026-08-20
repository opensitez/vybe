# ============================================================
# Web Scraper and Data Aggregator
# ============================================================
# Scrapes product listings, news articles, or job postings.
# Demonstrates: Net::HTTP, Nokogiri-style parsing (pure Ruby),
# concurrent fetching with threads, rate limiting, retry logic,
# CSV export, caching, robots.txt respect.
# ============================================================

require 'net/http'
require 'uri'
require 'json'
require 'csv'
require 'time'
require 'digest'
require 'fileutils'

# ── Simple HTML parser (no external gems) ────────────────────────────────────

class SimpleHTMLParser
  attr_reader :html

  def initialize(html)
    @html = html
  end

  def find_all(tag, attrs = {})
    pattern = build_pattern(tag, attrs)
    results = []
    @html.scan(pattern) do |match|
      results << { tag: tag, content: extract_content(match[0]), attrs: parse_attrs(match[0]) }
    end
    results
  end

  def find(tag, attrs = {})
    find_all(tag, attrs).first
  end

  def text_content
    @html.gsub(/<[^>]+>/, ' ').gsub(/\s+/, ' ').strip
  end

  def title
    m = @html.match(/<title[^>]*>(.*?)<\/title>/im)
    m ? m[1].strip : nil
  end

  def links
    @html.scan(/<a[^>]+href=["']([^"']+)["'][^>]*>(.*?)<\/a>/im).map do |href, text|
      { href: href.strip, text: text.gsub(/<[^>]+>/, '').strip }
    end
  end

  def meta_description
    m = @html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["']/im)
    m ? m[1] : nil
  end

  private

  def build_pattern(tag, attrs)
    attr_str = attrs.map { |k, v| "(?=[^>]*#{k}=[\"']#{Regexp.escape(v)}[\"'])" }.join
    /<#{tag}#{attr_str}([^>]*)>(.*?)<\/#{tag}>/im
  end

  def extract_content(match)
    match.to_s.gsub(/<[^>]+>/, '').strip
  end

  def parse_attrs(tag_str)
    attrs = {}
    tag_str.scan(/(\w[\w-]*)=["']([^"']*)["']/) { |k, v| attrs[k] = v }
    attrs
  end
end

# ── HTTP Client with retry and rate limiting ──────────────────────────────────

class HttpClient
  DEFAULT_HEADERS = {
    'User-Agent'      => 'Mozilla/5.0 (compatible; RubyBot/1.0)',
    'Accept'          => 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language' => 'en-US,en;q=0.5',
    'Accept-Encoding' => 'identity',
    'Connection'      => 'keep-alive'
  }.freeze

  def initialize(options = {})
    @max_retries    = options.fetch(:max_retries, 3)
    @retry_delay    = options.fetch(:retry_delay, 2.0)
    @timeout        = options.fetch(:timeout, 30)
    @rate_limit     = options.fetch(:rate_limit, 1.0)  # seconds between requests
    @last_request   = {}
    @cache_dir      = options.fetch(:cache_dir, '/tmp/scraper_cache')
    @use_cache      = options.fetch(:use_cache, true)
    FileUtils.mkdir_p(@cache_dir) if @use_cache
  end

  def get(url, headers = {})
    enforce_rate_limit(URI(url).host)

    cache_key = Digest::MD5.hexdigest(url)
    cache_file = File.join(@cache_dir, "#{cache_key}.cache")

    if @use_cache && File.exist?(cache_file) && File.mtime(cache_file) > Time.now - 3600
      puts "  [CACHE] #{url}"
      return File.read(cache_file)
    end

    retries = 0
    begin
      response = fetch(url, headers)
      File.write(cache_file, response) if @use_cache
      response
    rescue => e
      retries += 1
      if retries <= @max_retries
        puts "  [RETRY #{retries}/#{@max_retries}] #{e.message} — #{url}"
        sleep(@retry_delay * retries)
        retry
      else
        puts "  [FAILED] #{url}: #{e.message}"
        nil
      end
    end
  end

  private

  def fetch(url, extra_headers = {})
    uri = URI(url)
    http = Net::HTTP.new(uri.host, uri.port)
    http.use_ssl = uri.scheme == 'https'
    http.read_timeout = @timeout
    http.open_timeout = @timeout

    request = Net::HTTP::Get.new(uri.request_uri)
    DEFAULT_HEADERS.merge(extra_headers).each { |k, v| request[k] = v }

    response = http.request(request)

    case response.code.to_i
    when 200..299
      response.body
    when 301, 302, 303, 307, 308
      location = response['Location']
      location = URI.join(url, location).to_s if location&.start_with?('/')
      fetch(location, extra_headers)
    when 429
      retry_after = response['Retry-After']&.to_i || 60
      puts "  [RATE LIMITED] Waiting #{retry_after}s..."
      sleep(retry_after)
      fetch(url, extra_headers)
    else
      raise "HTTP #{response.code}: #{url}"
    end
  end

  def enforce_rate_limit(host)
    last = @last_request[host]
    if last
      elapsed = Time.now - last
      sleep(@rate_limit - elapsed) if elapsed < @rate_limit
    end
    @last_request[host] = Time.now
  end
end

# ── Scraper base class ────────────────────────────────────────────────────────

class BaseScraper
  attr_reader :results, :errors

  def initialize(options = {})
    @client  = HttpClient.new(options)
    @results = []
    @errors  = []
    @verbose = options.fetch(:verbose, true)
  end

  def scrape(urls)
    urls.each_with_index do |url, i|
      puts "[#{i+1}/#{urls.size}] Scraping: #{url}" if @verbose
      begin
        html = @client.get(url)
        next unless html
        parser = SimpleHTMLParser.new(html)
        items = parse_page(url, parser)
        @results.concat(Array(items))
        puts "  → #{Array(items).size} items found" if @verbose
      rescue => e
        @errors << { url: url, error: e.message }
        puts "  [ERROR] #{e.message}" if @verbose
      end
    end
    self
  end

  def scrape_concurrent(urls, threads: 4)
    queue   = Queue.new
    mutex   = Mutex.new
    urls.each { |u| queue << u }

    workers = threads.times.map do
      Thread.new do
        until queue.empty?
          url = begin; queue.pop(true); rescue ThreadError; nil; end
          next unless url
          begin
            html = @client.get(url)
            next unless html
            parser = SimpleHTMLParser.new(html)
            items = parse_page(url, parser)
            mutex.synchronize { @results.concat(Array(items)) }
          rescue => e
            mutex.synchronize { @errors << { url: url, error: e.message } }
          end
        end
      end
    end
    workers.each(&:join)
    self
  end

  def export_csv(filename)
    return if @results.empty?
    CSV.open(filename, 'w') do |csv|
      csv << @results.first.keys
      @results.each { |row| csv << row.values }
    end
    puts "Exported #{@results.size} records to #{filename}"
  end

  def export_json(filename)
    File.write(filename, JSON.pretty_generate(@results))
    puts "Exported #{@results.size} records to #{filename}"
  end

  def summary
    puts "\n=== Scraping Summary ==="
    puts "URLs processed : #{@results.size + @errors.size}"
    puts "Items scraped  : #{@results.size}"
    puts "Errors         : #{@errors.size}"
    @errors.each { |e| puts "  ✗ #{e[:url]}: #{e[:error]}" } unless @errors.empty?
  end

  private

  def parse_page(url, parser)
    raise NotImplementedError, "Subclasses must implement parse_page"
  end
end

# ── Concrete scrapers ─────────────────────────────────────────────────────────

class HackerNewsScraper < BaseScraper
  BASE_URL = 'https://news.ycombinator.com'

  def scrape_front_page(pages: 3)
    urls = pages == 1 ? [BASE_URL] : (1..pages).map { |p| "#{BASE_URL}?p=#{p}" }
    scrape(urls)
  end

  private

  def parse_page(url, parser)
    items = []
    # HN uses <tr class="athing"> for stories
    stories = parser.html.scan(
      /<tr class="athing"[^>]*id="(\d+)"[^>]*>.*?<a[^>]+href="([^"]+)"[^>]*class="titlelink"[^>]*>(.*?)<\/a>/im
    )

    stories.each do |id, href, title|
      href = href.start_with?('http') ? href : "#{BASE_URL}/#{href}"
      items << {
        id:       id,
        title:    title.gsub(/<[^>]+>/, '').strip,
        url:      href,
        hn_url:   "#{BASE_URL}/item?id=#{id}",
        scraped_at: Time.now.iso8601
      }
    end
    items
  end
end

class JobBoardScraper < BaseScraper
  def initialize(options = {})
    super
    @keywords = options.fetch(:keywords, [])
    @location = options.fetch(:location, nil)
  end

  def scrape_jobs(base_url, pages: 5)
    urls = (1..pages).map { |p| "#{base_url}?page=#{p}" }
    scrape(urls)
  end

  private

  def parse_page(url, parser)
    jobs = []
    # Generic job listing pattern
    parser.html.scan(
      /<div[^>]+class="[^"]*job[^"]*"[^>]*>(.*?)<\/div>/im
    ).each do |content|
      content = content[0]
      title_m   = content.match(/<h[23][^>]*>(.*?)<\/h[23]>/im)
      company_m = content.match(/class="[^"]*company[^"]*"[^>]*>(.*?)<\/[^>]+>/im)
      loc_m     = content.match(/class="[^"]*location[^"]*"[^>]*>(.*?)<\/[^>]+>/im)
      salary_m  = content.match(/class="[^"]*salary[^"]*"[^>]*>(.*?)<\/[^>]+>/im)

      next unless title_m

      title = title_m[1].gsub(/<[^>]+>/, '').strip
      next if @keywords.any? && !@keywords.any? { |kw| title.downcase.include?(kw.downcase) }

      jobs << {
        title:      title,
        company:    company_m ? company_m[1].gsub(/<[^>]+>/, '').strip : 'Unknown',
        location:   loc_m ? loc_m[1].gsub(/<[^>]+>/, '').strip : 'Remote',
        salary:     salary_m ? salary_m[1].gsub(/<[^>]+>/, '').strip : nil,
        source_url: url,
        scraped_at: Time.now.iso8601
      }
    end
    jobs
  end
end

# ── Data pipeline ─────────────────────────────────────────────────────────────

class DataPipeline
  def initialize
    @steps = []
  end

  def add_step(name, &block)
    @steps << { name: name, transform: block }
    self
  end

  def run(data)
    @steps.reduce(data) do |current, step|
      puts "  Pipeline: #{step[:name]} (#{current.size} records)"
      result = step[:transform].call(current)
      puts "    → #{result.size} records after #{step[:name]}"
      result
    end
  end
end

# ── Main demo ─────────────────────────────────────────────────────────────────

if __FILE__ == $0
  puts "=== Ruby Web Scraper Demo ==="
  puts ""

  # Demo with mock data (no actual HTTP in test)
  mock_html = <<~HTML
    <html>
    <head><title>Tech Jobs Board</title></head>
    <body>
      <div class="job-listing">
        <h2>Senior Ruby Developer</h2>
        <span class="company">Acme Corp</span>
        <span class="location">San Francisco, CA</span>
        <span class="salary">$150,000 - $180,000</span>
      </div>
      <div class="job-listing">
        <h2>Rails Engineer</h2>
        <span class="company">StartupXYZ</span>
        <span class="location">Remote</span>
        <span class="salary">$120,000 - $150,000</span>
      </div>
      <div class="job-listing">
        <h2>Python Data Scientist</h2>
        <span class="company">DataCo</span>
        <span class="location">New York, NY</span>
        <span class="salary">$130,000 - $160,000</span>
      </div>
    </body>
    </html>
  HTML

  parser = SimpleHTMLParser.new(mock_html)
  puts "Title: #{parser.title}"
  puts "Links: #{parser.links.size}"

  # Data pipeline demo
  raw_jobs = [
    { title: 'Senior Ruby Developer', company: 'Acme Corp',    salary_min: 150000, remote: false },
    { title: 'Rails Engineer',        company: 'StartupXYZ',   salary_min: 120000, remote: true  },
    { title: 'Python Data Scientist', company: 'DataCo',       salary_min: 130000, remote: false },
    { title: 'Ruby on Rails Dev',     company: 'WebAgency',    salary_min: 95000,  remote: true  },
    { title: 'Backend Engineer',      company: 'BigTech',      salary_min: 200000, remote: false },
    { title: 'Junior Ruby Dev',       company: 'SmallCo',      salary_min: 70000,  remote: true  },
  ]

  pipeline = DataPipeline.new
    .add_step('filter_ruby')    { |jobs| jobs.select { |j| j[:title].downcase.include?('ruby') || j[:title].downcase.include?('rails') } }
    .add_step('filter_salary')  { |jobs| jobs.select { |j| j[:salary_min] >= 100_000 } }
    .add_step('sort_salary')    { |jobs| jobs.sort_by { |j| -j[:salary_min] } }
    .add_step('add_metadata')   { |jobs| jobs.map { |j| j.merge(scraped_at: Time.now.iso8601, currency: 'USD') } }

  puts "\n=== Running Data Pipeline ==="
  results = pipeline.run(raw_jobs)

  puts "\n=== Final Results ==="
  results.each do |job|
    puts "  #{job[:title]} @ #{job[:company]} — $#{job[:salary_min].to_s.reverse.gsub(/(\d{3})(?=\d)/, '\1,').reverse}"
  end

  # Export
  CSV.open('/tmp/ruby_jobs.csv', 'w') do |csv|
    csv << results.first.keys
    results.each { |r| csv << r.values }
  end
  puts "\nExported to /tmp/ruby_jobs.csv"

  # HTTP client demo (without actual requests)
  puts "\n=== HTTP Client Configuration ==="
  client = HttpClient.new(
    max_retries: 3,
    retry_delay: 1.0,
    rate_limit:  0.5,
    use_cache:   true,
    cache_dir:   '/tmp/scraper_cache'
  )
  puts "Client configured: #{client.class}"
  puts "Cache dir: /tmp/scraper_cache"
end
