# ============================================================
# Lightweight ORM Framework (ActiveRecord-style)
# ============================================================
# A mini ORM with: model definition via DSL, associations
# (has_many, belongs_to, has_one), validations, callbacks,
# query builder, migrations, and connection pooling.
# Uses SQLite3 via pure Ruby (no gems needed for demo).
# ============================================================

require 'json'
require 'time'
require 'set'

# ── Connection pool ───────────────────────────────────────────────────────────

class ConnectionPool
  def initialize(size: 5, &connector)
    @size      = size
    @connector = connector
    @pool      = []
    @mutex     = Mutex.new
    @cond      = ConditionVariable.new
    @size.times { @pool << { conn: @connector.call, in_use: false } }
  end

  def with_connection
    conn = acquire
    begin
      yield conn
    ensure
      release(conn)
    end
  end

  def stats
    @mutex.synchronize do
      { total: @pool.size, in_use: @pool.count { |e| e[:in_use] }, available: @pool.count { |e| !e[:in_use] } }
    end
  end

  private

  def acquire
    @mutex.synchronize do
      loop do
        entry = @pool.find { |e| !e[:in_use] }
        if entry
          entry[:in_use] = true
          return entry[:conn]
        end
        @cond.wait(@mutex, 5)
        raise "Connection pool timeout" if @pool.none? { |e| !e[:in_use] }
      end
    end
  end

  def release(conn)
    @mutex.synchronize do
      entry = @pool.find { |e| e[:conn] == conn }
      entry[:in_use] = false if entry
      @cond.signal
    end
  end
end

# ── In-memory database (simulates SQLite) ────────────────────────────────────

class InMemoryDB
  def initialize
    @tables  = {}
    @seq     = Hash.new(0)
    @indices = {}
  end

  def create_table(name, columns)
    @tables[name.to_s] = []
    @indices[name.to_s] = {}
    columns
  end

  def insert(table, attrs)
    t = table.to_s
    @seq[t] += 1
    row = attrs.merge('id' => @seq[t], 'created_at' => Time.now.iso8601, 'updated_at' => Time.now.iso8601)
    @tables[t] << row
    row
  end

  def update(table, id, attrs)
    t = table.to_s
    row = @tables[t].find { |r| r['id'] == id }
    return nil unless row
    row.merge!(attrs.merge('updated_at' => Time.now.iso8601))
    row
  end

  def delete(table, id)
    t = table.to_s
    @tables[t].reject! { |r| r['id'] == id }
  end

  def find(table, id)
    @tables[table.to_s]&.find { |r| r['id'] == id }
  end

  def where(table, conditions = {}, order: nil, limit: nil, offset: 0)
    rows = @tables[table.to_s] || []
    rows = rows.select do |row|
      conditions.all? do |k, v|
        case v
        when Array  then v.include?(row[k.to_s])
        when Range  then v.include?(row[k.to_s])
        when Regexp then row[k.to_s].to_s.match?(v)
        else row[k.to_s] == v
        end
      end
    end
    rows = rows.sort_by { |r| order.map { |col, dir| dir == :desc ? -r[col.to_s].to_s.ord : r[col.to_s].to_s.ord } } if order
    rows = rows.drop(offset) if offset > 0
    rows = rows.first(limit) if limit
    rows
  end

  def count(table, conditions = {})
    where(table, conditions).size
  end

  def all(table)
    @tables[table.to_s]&.dup || []
  end

  def tables
    @tables.keys
  end
end

# Global DB instance
DB = InMemoryDB.new

# ── Query builder ─────────────────────────────────────────────────────────────

class QueryBuilder
  def initialize(model_class)
    @model_class = model_class
    @conditions  = {}
    @order       = nil
    @limit_val   = nil
    @offset_val  = 0
    @includes    = []
  end

  def where(conditions = {})
    @conditions.merge!(conditions)
    self
  end

  def order(*cols)
    @order = cols.map { |c| c.is_a?(Hash) ? c.first : [c, :asc] }
    self
  end

  def limit(n)
    @limit_val = n
    self
  end

  def offset(n)
    @offset_val = n
    self
  end

  def includes(*associations)
    @includes.concat(associations)
    self
  end

  def first(n = nil)
    results = execute
    n ? results.first(n) : results.first
  end

  def last(n = nil)
    results = execute
    n ? results.last(n) : results.last
  end

  def count
    DB.count(@model_class.table_name, @conditions)
  end

  def exists?
    count > 0
  end

  def pluck(*columns)
    execute.map { |r| columns.size == 1 ? r.send(columns.first) : columns.map { |c| r.send(c) } }
  end

  def to_a
    execute
  end

  def each(&block)
    execute.each(&block)
  end

  def map(&block)
    execute.map(&block)
  end

  def select(&block)
    execute.select(&block)
  end

  def inspect
    "#<QueryBuilder model=#{@model_class} conditions=#{@conditions} limit=#{@limit_val}>"
  end

  private

  def execute
    rows = DB.where(@model_class.table_name, @conditions,
                    order: @order, limit: @limit_val, offset: @offset_val)
    records = rows.map { |row| @model_class.instantiate(row) }

    # Eager loading
    @includes.each do |assoc|
      @model_class.eager_load(records, assoc)
    end

    records
  end
end

# ── Base model ────────────────────────────────────────────────────────────────

class Model
  # Class-level DSL
  class << self
    def table_name(name = nil)
      name ? @table_name = name : (@table_name || "#{self.name.downcase}s")
    end

    def column(name, type: :string, default: nil, null: true)
      @columns ||= {}
      @columns[name.to_sym] = { type: type, default: default, null: null }
      attr_accessor name.to_sym
    end

    def columns
      @columns || {}
    end

    # Associations
    def has_many(name, class_name: nil, foreign_key: nil)
      @associations ||= {}
      klass_name = class_name || name.to_s.chomp('s').capitalize
      fk         = foreign_key || "#{self.name.downcase}_id"
      @associations[name] = { type: :has_many, class_name: klass_name, foreign_key: fk }

      define_method(name) do
        klass = Object.const_get(klass_name)
        klass.where(fk => id)
      end
    end

    def belongs_to(name, class_name: nil, foreign_key: nil)
      @associations ||= {}
      klass_name = class_name || name.to_s.capitalize
      fk         = foreign_key || "#{name}_id"
      @associations[name] = { type: :belongs_to, class_name: klass_name, foreign_key: fk }

      define_method(name) do
        klass = Object.const_get(klass_name)
        fk_val = send(fk)
        fk_val ? klass.find(fk_val) : nil
      end

      define_method("#{name}=") do |record|
        send("#{fk}=", record&.id)
      end
    end

    def has_one(name, class_name: nil, foreign_key: nil)
      @associations ||= {}
      klass_name = class_name || name.to_s.capitalize
      fk         = foreign_key || "#{self.name.downcase}_id"
      @associations[name] = { type: :has_one, class_name: klass_name, foreign_key: fk }

      define_method(name) do
        klass = Object.const_get(klass_name)
        klass.where(fk => id).first
      end
    end

    def associations
      @associations || {}
    end

    # Validations
    def validates(field, **rules)
      @validations ||= []
      @validations << { field: field, rules: rules }
    end

    def validations
      @validations || []
    end

    # Callbacks
    %i[before_save after_save before_create after_create before_update after_update before_destroy after_destroy].each do |cb|
      define_method(cb) do |method_name = nil, &block|
        @callbacks ||= Hash.new { |h, k| h[k] = [] }
        @callbacks[cb] << (method_name || block)
      end

      define_method("run_#{cb}") do |instance|
        (@callbacks || {})[cb]&.each do |cb_item|
          cb_item.is_a?(Symbol) ? instance.send(cb_item) : cb_item.call(instance)
        end
      end
    end

    # Scopes
    def scope(name, &block)
      define_singleton_method(name) { |*args| block.call(*args) }
    end

    # Finders
    def find(id)
      row = DB.find(table_name, id)
      row ? instantiate(row) : nil
    end

    def find!(id)
      find(id) || raise("#{name} with id=#{id} not found")
    end

    def where(conditions = {})
      QueryBuilder.new(self).where(conditions)
    end

    def all
      QueryBuilder.new(self)
    end

    def first
      all.first
    end

    def last
      all.last
    end

    def count
      DB.count(table_name)
    end

    def create(attrs = {})
      record = new(attrs)
      record.save
      record
    end

    def create!(attrs = {})
      record = new(attrs)
      record.save!
      record
    end

    def instantiate(row)
      record = allocate
      record.instance_variable_set(:@persisted, true)
      record.instance_variable_set(:@attributes, row.dup)
      row.each { |k, v| record.send(:"#{k}=", v) if record.respond_to?(:"#{k}=") }
      record.instance_variable_set(:@id, row['id'])
      record
    end

    def eager_load(records, assoc_name)
      assoc = associations[assoc_name]
      return unless assoc

      case assoc[:type]
      when :has_many
        ids = records.map(&:id)
        klass = Object.const_get(assoc[:class_name])
        related = klass.where(assoc[:foreign_key] => ids).to_a
        grouped = related.group_by { |r| r.send(assoc[:foreign_key]) }
        records.each do |r|
          r.instance_variable_set(:"@#{assoc_name}_cache", grouped[r.id] || [])
        end
      end
    end

    def migrate
      DB.create_table(table_name, columns)
    end
  end

  attr_reader :id, :errors

  def initialize(attrs = {})
    @persisted  = false
    @attributes = {}
    @errors     = {}
    @changed    = Set.new

    # Set defaults
    self.class.columns.each do |name, opts|
      send(:"#{name}=", opts[:default]) unless opts[:default].nil?
    end

    attrs.each { |k, v| send(:"#{k}=", v) if respond_to?(:"#{k}=") }
  end

  def persisted?
    @persisted
  end

  def new_record?
    !@persisted
  end

  def valid?
    @errors = {}
    run_validations
    @errors.empty?
  end

  def save
    return false unless valid?
    run_callbacks(:before_save)
    if new_record?
      run_callbacks(:before_create)
      attrs = serializable_hash
      row   = DB.insert(self.class.table_name, attrs)
      @id   = row['id']
      send(:id=, @id) if respond_to?(:id=)
      @persisted = true
      run_callbacks(:after_create)
    else
      run_callbacks(:before_update)
      DB.update(self.class.table_name, @id, serializable_hash)
      run_callbacks(:after_update)
    end
    run_callbacks(:after_save)
    true
  end

  def save!
    save || raise("Validation failed: #{@errors}")
  end

  def update(attrs = {})
    attrs.each { |k, v| send(:"#{k}=", v) if respond_to?(:"#{k}=") }
    save
  end

  def destroy
    run_callbacks(:before_destroy)
    DB.delete(self.class.table_name, @id)
    @persisted = false
    run_callbacks(:after_destroy)
    self
  end

  def reload
    row = DB.find(self.class.table_name, @id)
    row&.each { |k, v| send(:"#{k}=", v) if respond_to?(:"#{k}=") }
    self
  end

  def to_h
    serializable_hash
  end

  def to_json(*args)
    to_h.to_json(*args)
  end

  def ==(other)
    other.is_a?(self.class) && id == other.id
  end

  def inspect
    attrs = serializable_hash.map { |k, v| "#{k}: #{v.inspect}" }.join(', ')
    "#<#{self.class.name} #{attrs}>"
  end

  private

  def serializable_hash
    self.class.columns.each_with_object({}) do |(name, _), h|
      h[name.to_s] = send(name) if respond_to?(name)
    end
  end

  def run_validations
    self.class.validations.each do |v|
      field = v[:field]
      value = respond_to?(field) ? send(field) : nil
      rules = v[:rules]

      if rules[:presence] && (value.nil? || value.to_s.strip.empty?)
        (@errors[field] ||= []) << "can't be blank"
      end
      if rules[:uniqueness] && value
        existing = self.class.where(field => value).first
        if existing && existing.id != @id
          (@errors[field] ||= []) << "has already been taken"
        end
      end
      if rules[:format] && value && !value.to_s.match?(rules[:format])
        (@errors[field] ||= []) << "is invalid"
      end
      if rules[:length] && value
        len = value.to_s.length
        (@errors[field] ||= []) << "is too short (min #{rules[:length][:min]})" if rules[:length][:min] && len < rules[:length][:min]
        (@errors[field] ||= []) << "is too long (max #{rules[:length][:max]})"  if rules[:length][:max] && len > rules[:length][:max]
      end
      if rules[:numericality] && value
        (@errors[field] ||= []) << "must be a number" unless value.is_a?(Numeric)
        (@errors[field] ||= []) << "must be > #{rules[:numericality][:greater_than]}" if rules[:numericality][:greater_than] && value <= rules[:numericality][:greater_than]
      end
    end
  end

  def run_callbacks(event)
    self.class.send(:"run_#{event}", self)
  end
end

# ── Concrete models ───────────────────────────────────────────────────────────

class User < Model
  table_name 'users'

  column :name,       type: :string
  column :email,      type: :string
  column :role,       type: :string, default: 'user'
  column :active,     type: :boolean, default: true
  column :age,        type: :integer
  column :created_at, type: :datetime
  column :updated_at, type: :datetime

  validates :name,  presence: true, length: { min: 2, max: 100 }
  validates :email, presence: true, format: /\A[^@\s]+@[^@\s]+\z/, uniqueness: true
  validates :age,   numericality: { greater_than: 0 }

  has_many :posts
  has_many :orders, class_name: 'Order'
  has_one  :profile

  before_create { |u| u.email = u.email&.downcase }
  after_create  { |u| puts "  [Callback] User created: #{u.name}" }

  scope :active,  -> { where(active: true) }
  scope :admins,  -> { where(role: 'admin') }
  scope :by_name, ->(name) { where(name: name) }

  def full_display
    "#{name} <#{email}> (#{role})"
  end

  def admin?
    role == 'admin'
  end
end

class Post < Model
  table_name 'posts'

  column :title,      type: :string
  column :body,       type: :text
  column :user_id,    type: :integer
  column :published,  type: :boolean, default: false
  column :views,      type: :integer, default: 0
  column :created_at, type: :datetime
  column :updated_at, type: :datetime

  validates :title, presence: true, length: { min: 5, max: 200 }
  validates :body,  presence: true

  belongs_to :user

  scope :published, -> { where(published: true) }
  scope :popular,   -> { where(views: (100..Float::INFINITY)) }

  before_save { |p| p.title = p.title&.strip }
end

class Order < Model
  table_name 'orders'

  column :user_id,    type: :integer
  column :total,      type: :decimal
  column :status,     type: :string, default: 'pending'
  column :created_at, type: :datetime
  column :updated_at, type: :datetime

  validates :total,  numericality: { greater_than: 0 }
  validates :status, presence: true

  belongs_to :user
end

# ── Demo ──────────────────────────────────────────────────────────────────────

if __FILE__ == $0
  puts "=== ORM Framework Demo ==="
  puts ""

  # Run migrations
  [User, Post, Order].each(&:migrate)
  puts "Migrations complete: #{DB.tables.join(', ')}"
  puts ""

  # Create users
  puts "=== Creating Users ==="
  alice = User.create!(name: 'Alice Smith', email: 'alice@example.com', role: 'admin', age: 30)
  bob   = User.create!(name: 'Bob Jones',   email: 'bob@example.com',   role: 'user',  age: 25)
  carol = User.create!(name: 'Carol White', email: 'carol@example.com', role: 'user',  age: 35)

  puts "Created: #{alice.full_display}"
  puts "Created: #{bob.full_display}"

  # Validation failure
  puts "\n=== Validation Demo ==="
  bad_user = User.new(name: 'X', email: 'not-an-email', age: -1)
  puts "Valid? #{bad_user.valid?}"
  puts "Errors: #{bad_user.errors}"

  # Uniqueness validation
  dup_user = User.new(name: 'Alice Duplicate', email: 'alice@example.com', age: 20)
  puts "Duplicate email valid? #{dup_user.valid?}"
  puts "Errors: #{dup_user.errors}"

  # Create posts
  puts "\n=== Creating Posts ==="
  p1 = Post.create!(title: 'Introduction to Ruby', body: 'Ruby is a dynamic language...', user_id: alice.id, published: true, views: 250)
  p2 = Post.create!(title: 'Advanced Metaprogramming', body: 'Ruby metaprogramming allows...', user_id: alice.id, published: true, views: 50)
  p3 = Post.create!(title: 'Getting Started with Rails', body: 'Rails is a web framework...', user_id: bob.id, published: false, views: 10)

  # Create orders
  Order.create!(user_id: alice.id, total: 99.99,  status: 'completed')
  Order.create!(user_id: alice.id, total: 149.50, status: 'completed')
  Order.create!(user_id: bob.id,   total: 29.99,  status: 'pending')

  # Queries
  puts "\n=== Query Builder ==="
  puts "All users: #{User.count}"
  puts "Active users: #{User.active.count}"
  puts "Admins: #{User.admins.pluck(:name)}"

  puts "\nPublished posts: #{Post.published.count}"
  puts "Popular posts: #{Post.published.select { |p| p.views >= 100 }.map(&:title)}"

  # Associations
  puts "\n=== Associations ==="
  alice_posts = alice.posts.to_a
  puts "Alice's posts: #{alice_posts.map(&:title)}"

  alice_orders = alice.orders.to_a
  puts "Alice's orders: #{alice_orders.map { |o| "$#{o.total}" }}"

  post_author = p1.user
  puts "Post '#{p1.title}' author: #{post_author.name}"

  # Update
  puts "\n=== Update ==="
  p3.update(published: true, views: 100)
  puts "Post updated: #{p3.title} published=#{p3.published}"

  # Chained queries
  puts "\n=== Chained Queries ==="
  recent_published = Post.published.limit(2).to_a
  puts "Recent published (limit 2): #{recent_published.map(&:title)}"

  # Destroy
  puts "\n=== Destroy ==="
  temp = User.create!(name: 'Temp User', email: 'temp@example.com', age: 20)
  puts "Before destroy: #{User.count} users"
  temp.destroy
  puts "After destroy: #{User.count} users"

  # Serialization
  puts "\n=== Serialization ==="
  puts "Alice as JSON: #{alice.to_json}"
end
