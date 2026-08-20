# ============================================================
# Data Processing Pipeline
# ============================================================
# ETL pipeline: extract from CSV/JSON, transform with
# validations and enrichment, load to multiple targets.
# Demonstrates Ruby's strengths: blocks, Enumerable,
# method chaining, lazy enumerators, struct, comparable.
# ============================================================

require 'csv'
require 'json'
require 'date'
require 'set'

# ── Domain models ─────────────────────────────────────────────────────────────

Customer = Struct.new(:id, :name, :email, :country, :signup_date, :plan, keyword_init: true) do
  def active?
    !signup_date.nil?
  end

  def enterprise?
    plan == 'enterprise'
  end

  def days_since_signup
    (Date.today - signup_date).to_i
  end

  def to_h
    super.transform_values { |v| v.is_a?(Date) ? v.iso8601 : v }
  end
end

Order = Struct.new(:id, :customer_id, :amount, :currency, :status, :created_at, :items, keyword_init: true) do
  include Comparable

  def <=>(other)
    amount <=> other.amount
  end

  def completed?
    status == 'completed'
  end

  def refunded?
    status == 'refunded'
  end

  def amount_usd
    case currency
    when 'EUR' then amount * 1.08
    when 'GBP' then amount * 1.27
    when 'CAD' then amount * 0.74
    else amount
    end
  end
end

# ── Validation framework ──────────────────────────────────────────────────────

module Validatable
  def self.included(base)
    base.instance_variable_set(:@validations, [])
    base.extend(ClassMethods)
  end

  module ClassMethods
    def validates(field, **rules)
      @validations << { field: field, rules: rules }
    end

    def validations
      @validations
    end
  end

  def valid?
    errors.empty?
  end

  def errors
    @errors ||= validate
  end

  private

  def validate
    errs = {}
    self.class.validations.each do |v|
      field = v[:field]
      value = respond_to?(field) ? send(field) : nil
      rules = v[:rules]

      if rules[:presence] && (value.nil? || value.to_s.strip.empty?)
        (errs[field] ||= []) << "can't be blank"
      end

      if rules[:format] && value && !value.to_s.match?(rules[:format])
        (errs[field] ||= []) << "is invalid"
      end

      if rules[:inclusion] && value && !rules[:inclusion].include?(value)
        (errs[field] ||= []) << "is not included in the list"
      end

      if rules[:numericality] && value
        (errs[field] ||= []) << "must be greater than 0" if rules[:numericality][:greater_than] && value <= rules[:numericality][:greater_than]
        (errs[field] ||= []) << "must be less than #{rules[:numericality][:less_than]}" if rules[:numericality][:less_than] && value >= rules[:numericality][:less_than]
      end
    end
    errs
  end
end

class CustomerRecord
  include Validatable

  attr_accessor :id, :name, :email, :country, :plan

  validates :name,    presence: true
  validates :email,   presence: true, format: /\A[^@\s]+@[^@\s]+\z/
  validates :country, presence: true
  validates :plan,    inclusion: %w[free starter pro enterprise]

  def initialize(attrs = {})
    attrs.each { |k, v| send(:"#{k}=", v) if respond_to?(:"#{k}=") }
  end
end

# ── Transformers ──────────────────────────────────────────────────────────────

module Transformers
  def self.normalize_email(email)
    return nil unless email
    email.downcase.strip
  end

  def self.normalize_country(country)
    COUNTRY_CODES[country.upcase] || country
  end

  def self.parse_date(str)
    return nil unless str
    Date.parse(str)
  rescue Date::Error
    nil
  end

  def self.sanitize_name(name)
    return nil unless name
    name.strip.split.map(&:capitalize).join(' ')
  end

  COUNTRY_CODES = {
    'US' => 'United States', 'GB' => 'United Kingdom',
    'CA' => 'Canada',        'AU' => 'Australia',
    'DE' => 'Germany',       'FR' => 'France',
    'JP' => 'Japan',         'IN' => 'India',
    'BR' => 'Brazil',        'MX' => 'Mexico'
  }.freeze
end

# ── ETL Pipeline ──────────────────────────────────────────────────────────────

class Pipeline
  Result = Struct.new(:data, :errors, :stats, keyword_init: true)

  def initialize
    @steps   = []
    @filters = []
    @stats   = Hash.new(0)
  end

  def extract(&block)
    @extractor = block
    self
  end

  def transform(name, &block)
    @steps << { name: name, fn: block }
    self
  end

  def filter(name, &block)
    @filters << { name: name, fn: block }
    self
  end

  def load(&block)
    @loader = block
    self
  end

  def run
    puts "=== Pipeline Starting ==="
    start = Time.now

    # Extract
    raw = @extractor.call
    @stats[:extracted] = raw.size
    puts "Extracted: #{raw.size} records"

    # Transform
    transformed = raw.map.with_index do |record, i|
      @steps.reduce(record) do |r, step|
        begin
          step[:fn].call(r)
        rescue => e
          @stats[:transform_errors] += 1
          puts "  Transform error on record #{i}: #{e.message}"
          nil
        end
      end
    end.compact

    @stats[:transformed] = transformed.size
    puts "Transformed: #{transformed.size} records"

    # Filter
    filtered = transformed.select do |record|
      @filters.all? { |f| f[:fn].call(record) }
    end

    @stats[:filtered_out] = transformed.size - filtered.size
    @stats[:passed_filters] = filtered.size
    puts "After filters: #{filtered.size} records (#{@stats[:filtered_out]} filtered out)"

    # Load
    errors = []
    if @loader
      filtered.each do |record|
        begin
          @loader.call(record)
          @stats[:loaded] += 1
        rescue => e
          errors << { record: record, error: e.message }
          @stats[:load_errors] += 1
        end
      end
    end

    elapsed = Time.now - start
    @stats[:elapsed_ms] = (elapsed * 1000).round(1)
    puts "Loaded: #{@stats[:loaded]} records in #{@stats[:elapsed_ms]}ms"
    puts "=== Pipeline Complete ==="

    Result.new(data: filtered, errors: errors, stats: @stats.dup)
  end
end

# ── Analytics ─────────────────────────────────────────────────────────────────

class Analytics
  def initialize(orders)
    @orders = orders
  end

  def revenue_by_country(customers)
    customer_map = customers.each_with_object({}) { |c, h| h[c.id] = c }

    @orders
      .select(&:completed?)
      .group_by { |o| customer_map[o.customer_id]&.country || 'Unknown' }
      .transform_values { |orders| orders.sum(&:amount_usd) }
      .sort_by { |_, v| -v }
      .to_h
  end

  def revenue_by_month
    @orders
      .select(&:completed?)
      .group_by { |o| o.created_at.strftime('%Y-%m') }
      .transform_values { |orders| orders.sum(&:amount_usd) }
      .sort
      .to_h
  end

  def top_customers(n: 10)
    @orders
      .select(&:completed?)
      .group_by(&:customer_id)
      .transform_values { |orders| orders.sum(&:amount_usd) }
      .sort_by { |_, v| -v }
      .first(n)
  end

  def cohort_analysis(customers)
    customers
      .group_by { |c| c.signup_date.strftime('%Y-%m') }
      .transform_values do |cohort|
        cohort_ids = cohort.map(&:id).to_set
        cohort_orders = @orders.select { |o| cohort_ids.include?(o.customer_id) && o.completed? }
        {
          customers:   cohort.size,
          orders:      cohort_orders.size,
          revenue:     cohort_orders.sum(&:amount_usd).round(2),
          avg_revenue: cohort.size > 0 ? (cohort_orders.sum(&:amount_usd) / cohort.size).round(2) : 0
        }
      end
      .sort
      .to_h
  end

  def summary_stats
    completed = @orders.select(&:completed?)
    amounts   = completed.map(&:amount_usd)
    return {} if amounts.empty?

    sorted = amounts.sort
    n      = sorted.size
    mean   = amounts.sum / n
    median = n.odd? ? sorted[n/2] : (sorted[n/2 - 1] + sorted[n/2]) / 2.0
    variance = amounts.sum { |a| (a - mean) ** 2 } / n
    std_dev  = Math.sqrt(variance)

    {
      count:    n,
      total:    amounts.sum.round(2),
      mean:     mean.round(2),
      median:   median.round(2),
      std_dev:  std_dev.round(2),
      min:      sorted.first.round(2),
      max:      sorted.last.round(2),
      p25:      sorted[(n * 0.25).floor].round(2),
      p75:      sorted[(n * 0.75).floor].round(2),
      p95:      sorted[(n * 0.95).floor].round(2)
    }
  end
end

# ── Lazy enumerator demo ──────────────────────────────────────────────────────

class InfiniteSequence
  include Enumerable

  def initialize(&generator)
    @generator = generator
  end

  def each
    n = 0
    loop { yield @generator.call(n); n += 1 }
  end

  def take_while_lazy(&block)
    lazy.take_while(&block).to_a
  end
end

# ── Main demo ─────────────────────────────────────────────────────────────────

if __FILE__ == $0
  puts "=== Data Processing Pipeline Demo ==="
  puts ""

  # Generate sample data
  plans    = %w[free starter pro enterprise]
  countries = %w[US GB CA AU DE FR JP IN BR MX]

  customers = 50.times.map do |i|
    Customer.new(
      id:          i + 1,
      name:        "Customer #{i + 1}",
      email:       "customer#{i+1}@example.com",
      country:     countries[i % countries.size],
      signup_date: Date.today - rand(365 * 3),
      plan:        plans[i % plans.size]
    )
  end

  orders = 200.times.map do |i|
    Order.new(
      id:          i + 1,
      customer_id: rand(1..50),
      amount:      (rand * 500 + 10).round(2),
      currency:    %w[USD EUR GBP CAD][rand(4)],
      status:      %w[completed completed completed refunded pending][rand(5)],
      created_at:  Date.today - rand(365),
      items:       rand(1..5)
    )
  end

  # Validation demo
  puts "=== Validation Demo ==="
  valid_rec = CustomerRecord.new(name: 'Alice Smith', email: 'alice@example.com', country: 'US', plan: 'pro')
  bad_rec   = CustomerRecord.new(name: '', email: 'not-an-email', country: 'US', plan: 'invalid')

  puts "Valid record: #{valid_rec.valid?}"
  puts "Invalid record errors: #{bad_rec.errors}"

  # Pipeline demo
  puts "\n=== ETL Pipeline ==="
  output = []

  result = Pipeline.new
    .extract { customers }
    .transform('normalize') do |c|
      Customer.new(
        id:          c.id,
        name:        Transformers.sanitize_name(c.name),
        email:       Transformers.normalize_email(c.email),
        country:     Transformers.normalize_country(c.country),
        signup_date: c.signup_date,
        plan:        c.plan
      )
    end
    .filter('active')      { |c| c.active? }
    .filter('valid_email') { |c| c.email&.include?('@') }
    .load { |c| output << c }
    .run

  puts "\nPipeline stats: #{result.stats}"

  # Analytics demo
  puts "\n=== Analytics ==="
  analytics = Analytics.new(orders)

  stats = analytics.summary_stats
  puts "Order Statistics:"
  stats.each { |k, v| puts "  #{k.to_s.ljust(10)}: #{v}" }

  puts "\nRevenue by Country (top 5):"
  analytics.revenue_by_country(customers).first(5).each do |country, rev|
    puts "  #{country.ljust(15)}: $#{rev.round(2)}"
  end

  puts "\nTop 5 Customers by Revenue:"
  analytics.top_customers(n: 5).each do |cust_id, rev|
    puts "  Customer #{cust_id}: $#{rev.round(2)}"
  end

  puts "\nCohort Analysis (last 3 months):"
  analytics.cohort_analysis(customers).last(3).each do |month, data|
    puts "  #{month}: #{data[:customers]} customers, $#{data[:avg_revenue]} avg revenue"
  end

  # Lazy enumerator demo
  puts "\n=== Lazy Enumerators ==="
  fibonacci = InfiniteSequence.new do |n|
    a, b = 0, 1
    n.times { a, b = b, a + b }
    a
  end

  fibs_under_1000 = fibonacci.lazy.select { |n| n.even? }.take_while { |n| n < 1000 }.to_a
  puts "Even Fibonacci numbers < 1000: #{fibs_under_1000}"

  primes = InfiniteSequence.new do |n|
    # nth prime (simple sieve)
    count = 0
    num   = 1
    loop do
      num += 1
      count += 1 if (2...num).none? { |d| num % d == 0 }
      break num if count > n
    end
  end

  puts "First 10 primes: #{primes.lazy.first(10).to_a}"

  # Enumerable power demo
  puts "\n=== Enumerable Power ==="
  order_stats = orders
    .select(&:completed?)
    .group_by { |o| o.currency }
    .transform_values { |os| { count: os.size, total: os.sum(&:amount).round(2), avg: (os.sum(&:amount) / os.size).round(2) } }

  puts "Orders by currency:"
  order_stats.each { |cur, s| puts "  #{cur}: #{s[:count]} orders, total #{s[:total]}, avg #{s[:avg]}" }
end
