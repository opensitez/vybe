# ============================================================
# Background Job Scheduler and Worker System
# ============================================================
# A Sidekiq/Resque-style job queue with priorities, retries,
# scheduling, worker pools, and monitoring. Pure Ruby with
# threads — no external dependencies.
# ============================================================

require 'thread'
require 'json'
require 'time'
require 'logger'
require 'digest'

# ── Job definition ────────────────────────────────────────────────────────────

class Job
  PRIORITIES = { critical: 0, high: 1, normal: 2, low: 3 }.freeze

  attr_reader :id, :queue, :worker_class, :args, :priority,
              :created_at, :scheduled_at, :attempts, :max_attempts,
              :error_history

  attr_accessor :status, :started_at, :completed_at, :result, :last_error

  def initialize(worker_class:, args: [], queue: :default, priority: :normal,
                 max_attempts: 3, scheduled_at: nil)
    @id            = Digest::SHA1.hexdigest("#{worker_class}#{args}#{Time.now.to_f}#{rand}")[0..15]
    @worker_class  = worker_class
    @args          = args
    @queue         = queue
    @priority      = priority
    @max_attempts  = max_attempts
    @created_at    = Time.now
    @scheduled_at  = scheduled_at || Time.now
    @attempts      = 0
    @status        = :pending
    @error_history = []
  end

  def ready?
    @scheduled_at <= Time.now
  end

  def retryable?
    @attempts < @max_attempts
  end

  def priority_value
    PRIORITIES.fetch(@priority, 2)
  end

  def duration
    return nil unless @started_at && @completed_at
    @completed_at - @started_at
  end

  def to_h
    {
      id:            @id,
      worker_class:  @worker_class,
      queue:         @queue,
      priority:      @priority,
      status:        @status,
      attempts:      @attempts,
      max_attempts:  @max_attempts,
      created_at:    @created_at.iso8601,
      scheduled_at:  @scheduled_at.iso8601,
      started_at:    @started_at&.iso8601,
      completed_at:  @completed_at&.iso8601,
      duration_ms:   duration ? (duration * 1000).round(1) : nil,
      last_error:    @last_error
    }
  end
end

# ── Worker base class ─────────────────────────────────────────────────────────

class BaseWorker
  class << self
    def queue(name = nil)
      name ? @queue = name : (@queue || :default)
    end

    def priority(level = nil)
      level ? @priority = level : (@priority || :normal)
    end

    def max_attempts(n = nil)
      n ? @max_attempts = n : (@max_attempts || 3)
    end

    def retry_in(seconds = nil)
      seconds ? @retry_in = seconds : (@retry_in || 60)
    end

    def perform_async(*args)
      Scheduler.instance.enqueue(
        worker_class: name,
        args:         args,
        queue:        queue,
        priority:     priority,
        max_attempts: max_attempts
      )
    end

    def perform_in(delay, *args)
      Scheduler.instance.enqueue(
        worker_class: name,
        args:         args,
        queue:        queue,
        priority:     priority,
        max_attempts: max_attempts,
        scheduled_at: Time.now + delay
      )
    end

    def perform_at(time, *args)
      Scheduler.instance.enqueue(
        worker_class: name,
        args:         args,
        queue:        queue,
        priority:     priority,
        max_attempts: max_attempts,
        scheduled_at: time
      )
    end
  end

  def perform(*args)
    raise NotImplementedError, "#{self.class}#perform must be implemented"
  end

  def logger
    Scheduler.instance.logger
  end
end

# ── Priority queue ────────────────────────────────────────────────────────────

class PriorityQueue
  def initialize
    @heap  = []
    @mutex = Mutex.new
  end

  def push(job)
    @mutex.synchronize do
      @heap << job
      @heap.sort_by! { |j| [j.priority_value, j.scheduled_at] }
    end
  end

  def pop
    @mutex.synchronize do
      idx = @heap.index { |j| j.ready? }
      idx ? @heap.delete_at(idx) : nil
    end
  end

  def peek
    @mutex.synchronize { @heap.find(&:ready?) }
  end

  def size
    @mutex.synchronize { @heap.size }
  end

  def pending_count
    @mutex.synchronize { @heap.count { |j| j.status == :pending } }
  end

  def scheduled_count
    @mutex.synchronize { @heap.count { |j| !j.ready? } }
  end

  def all
    @mutex.synchronize { @heap.dup }
  end
end

# ── Scheduler (singleton) ─────────────────────────────────────────────────────

class Scheduler
  include Comparable

  attr_reader :logger, :stats

  @instance = nil
  @mutex    = Mutex.new

  def self.instance
    @mutex.synchronize { @instance ||= new }
  end

  def initialize
    @queues      = Hash.new { |h, k| h[k] = PriorityQueue.new }
    @workers     = {}
    @job_history = []
    @history_max = 1000
    @running     = false
    @threads     = []
    @mutex       = Mutex.new
    @cond        = ConditionVariable.new
    @logger      = Logger.new($stdout)
    @logger.formatter = proc { |sev, time, _, msg| "[#{time.strftime('%H:%M:%S')}] #{sev}: #{msg}\n" }
    @stats       = Hash.new(0)
    @middleware  = []
    @callbacks   = Hash.new { |h, k| h[k] = [] }
  end

  def enqueue(**opts)
    job = Job.new(**opts)
    @queues[job.queue].push(job)
    @stats[:enqueued] += 1
    trigger(:job_enqueued, job)
    logger.debug("Enqueued #{job.worker_class}##{job.id} → #{job.queue}")
    job.id
  end

  def start(workers: 4, queues: [:default])
    @running = true
    @watched_queues = queues

    workers.times do |i|
      thread = Thread.new do
        Thread.current.name = "worker-#{i}"
        worker_loop
      end
      @threads << thread
    end

    # Scheduler thread for delayed jobs
    @scheduler_thread = Thread.new do
      Thread.current.name = "scheduler"
      scheduler_loop
    end

    logger.info("Scheduler started: #{workers} workers, queues: #{queues.join(', ')}")
    self
  end

  def stop(timeout: 30)
    @running = false
    @threads.each { |t| t.join(timeout) }
    @scheduler_thread&.join(5)
    logger.info("Scheduler stopped. Stats: #{@stats}")
  end

  def use(middleware_class, **opts)
    @middleware << middleware_class.new(**opts)
    self
  end

  def on(event, &block)
    @callbacks[event] << block
    self
  end

  def queue_stats
    @queues.transform_values do |q|
      { pending: q.pending_count, scheduled: q.scheduled_count, total: q.size }
    end
  end

  def job_history(limit: 50, status: nil)
    @mutex.synchronize do
      jobs = status ? @job_history.select { |j| j.status == status } : @job_history
      jobs.last(limit)
    end
  end

  def find_job(id)
    @mutex.synchronize do
      @queues.values.flat_map(&:all).find { |j| j.id == id } ||
        @job_history.find { |j| j.id == id }
    end
  end

  private

  def worker_loop
    while @running
      job = dequeue
      if job
        execute(job)
      else
        sleep(0.1)
      end
    end
  end

  def scheduler_loop
    while @running
      sleep(1)
      # Re-enqueue scheduled jobs that are now ready
      @queues.each_value do |queue|
        queue.all.select { |j| j.status == :pending && j.ready? }.each do |job|
          # Already in queue, just needs to be picked up
        end
      end
    end
  end

  def dequeue
    @watched_queues.each do |queue_name|
      job = @queues[queue_name].pop
      return job if job
    end
    nil
  end

  def execute(job)
    job.status     = :running
    job.started_at = Time.now
    job.attempts  += 1
    @stats[:started] += 1
    trigger(:job_started, job)

    begin
      worker = Object.const_get(job.worker_class).new
      run_with_middleware(job) { worker.perform(*job.args) }

      job.status       = :completed
      job.completed_at = Time.now
      @stats[:completed] += 1
      trigger(:job_completed, job)
      logger.info("✓ #{job.worker_class}##{job.id} (#{(job.duration * 1000).round}ms)")

    rescue => e
      job.last_error = "#{e.class}: #{e.message}"
      job.error_history << { error: job.last_error, at: Time.now.iso8601, attempt: job.attempts }
      @stats[:failed] += 1
      trigger(:job_failed, job, e)
      logger.error("✗ #{job.worker_class}##{job.id}: #{e.message}")

      if job.retryable?
        delay = 2 ** job.attempts  # exponential backoff
        job.status       = :pending
        job.scheduled_at = Time.now + delay
        @queues[job.queue].push(job)
        @stats[:retried] += 1
        logger.info("  Retry #{job.attempts}/#{job.max_attempts} in #{delay}s")
      else
        job.status       = :dead
        job.completed_at = Time.now
        @stats[:dead] += 1
        logger.error("  Job #{job.id} exhausted retries — moved to dead queue")
      end
    ensure
      @mutex.synchronize do
        @job_history << job
        @job_history.shift if @job_history.size > @history_max
      end
    end
  end

  def run_with_middleware(job, &block)
    chain = @middleware.reduce(block) do |inner, mw|
      -> { mw.call(job, inner) }
    end
    chain.call
  end

  def trigger(event, *args)
    @callbacks[event].each { |cb| cb.call(*args) rescue nil }
  end
end

# ── Middleware ────────────────────────────────────────────────────────────────

class TimingMiddleware
  def call(job, next_fn)
    start = Time.now
    next_fn.call
    elapsed = Time.now - start
    Scheduler.instance.logger.debug("  Timing: #{(elapsed * 1000).round(1)}ms")
  end
end

class LoggingMiddleware
  def call(job, next_fn)
    Scheduler.instance.logger.debug("  Args: #{job.args.inspect}")
    next_fn.call
  end
end

# ── Concrete workers ──────────────────────────────────────────────────────────

class EmailWorker < BaseWorker
  queue    :email
  priority :high
  max_attempts 5

  def perform(to:, subject:, body:, **opts)
    logger.info("  Sending email to #{to}: #{subject}")
    sleep(0.05)  # simulate SMTP
    raise "SMTP timeout" if rand < 0.05  # 5% failure rate
    logger.info("  Email sent to #{to}")
    { delivered: true, message_id: SecureRandom.hex(8) rescue rand(10000).to_s }
  end
end

class ReportWorker < BaseWorker
  queue    :reports
  priority :normal
  max_attempts 2

  def perform(report_type:, user_id:, params: {})
    logger.info("  Generating #{report_type} report for user #{user_id}")
    sleep(0.1 + rand * 0.2)  # simulate report generation
    rows = rand(100..10000)
    logger.info("  Report complete: #{rows} rows")
    { rows: rows, format: 'csv', size_kb: rows / 10 }
  end
end

class DataSyncWorker < BaseWorker
  queue    :sync
  priority :low
  max_attempts 3

  def perform(source:, destination:, table:)
    logger.info("  Syncing #{table}: #{source} → #{destination}")
    sleep(0.2 + rand * 0.3)
    records = rand(1000..50000)
    logger.info("  Synced #{records} records")
    { records_synced: records, duration_ms: rand(200..500) }
  end
end

class CriticalAlertWorker < BaseWorker
  queue    :critical
  priority :critical
  max_attempts 10

  def perform(alert_type:, severity:, message:)
    logger.info("  🚨 ALERT [#{severity}] #{alert_type}: #{message}")
    sleep(0.01)
    { notified: true, channels: %w[pagerduty slack email] }
  end
end

# ── Recurring job scheduler ───────────────────────────────────────────────────

class RecurringSchedule
  def initialize(scheduler)
    @scheduler = scheduler
    @jobs      = []
    @running   = false
  end

  def every(interval, worker_class, *args, **kwargs)
    @jobs << { interval: interval, worker: worker_class, args: args, kwargs: kwargs, last_run: nil }
    self
  end

  def cron(expression, worker_class, *args)
    # Simplified cron: just store for demo
    @jobs << { cron: expression, worker: worker_class, args: args, last_run: nil }
    self
  end

  def start
    @running = true
    @thread  = Thread.new do
      while @running
        @jobs.each do |job|
          next unless job[:interval]
          if job[:last_run].nil? || Time.now - job[:last_run] >= job[:interval]
            job[:last_run] = Time.now
            @scheduler.enqueue(
              worker_class: job[:worker],
              args:         job[:args],
              queue:        :default
            )
          end
        end
        sleep(1)
      end
    end
    self
  end

  def stop
    @running = false
    @thread&.join(5)
  end
end

# ── Demo ──────────────────────────────────────────────────────────────────────

if __FILE__ == $0
  puts "=== Job Scheduler Demo ==="
  puts ""

  scheduler = Scheduler.instance

  # Middleware
  scheduler.use(TimingMiddleware)
  scheduler.use(LoggingMiddleware)

  # Callbacks
  completed_count = 0
  failed_count    = 0

  scheduler.on(:job_completed) { |job| completed_count += 1 }
  scheduler.on(:job_failed)    { |job, err| failed_count += 1 }

  # Start workers
  scheduler.start(workers: 3, queues: [:critical, :email, :reports, :sync, :default])

  # Enqueue jobs
  puts "Enqueuing jobs..."

  # Critical alert
  scheduler.enqueue(
    worker_class: 'CriticalAlertWorker',
    args:         [],
    queue:        :critical,
    priority:     :critical,
    max_attempts: 10
  )

  # Batch of emails
  5.times do |i|
    EmailWorker.perform_async(
      to:      "user#{i}@example.com",
      subject: "Welcome to our service",
      body:    "Hello User #{i}!"
    )
  end

  # Reports
  3.times do |i|
    ReportWorker.perform_async(
      report_type: %w[sales inventory users][i],
      user_id:     i + 1
    )
  end

  # Delayed job
  DataSyncWorker.perform_in(2,
    source:      'production_db',
    destination: 'analytics_db',
    table:       'orders'
  )

  # Scheduled job
  DataSyncWorker.perform_at(Time.now + 3,
    source:      'production_db',
    destination: 'warehouse',
    table:       'customers'
  )

  puts "Jobs enqueued. Processing..."
  sleep(3)

  # Stats
  puts "\n=== Scheduler Statistics ==="
  puts "Stats: #{scheduler.stats}"
  puts "\nQueue Status:"
  scheduler.queue_stats.each do |queue, stats|
    puts "  #{queue}: #{stats}"
  end

  puts "\nRecent Job History (last 5):"
  scheduler.job_history(limit: 5).each do |job|
    status_icon = { completed: '✓', failed: '✗', dead: '💀', running: '⟳', pending: '⏳' }[job.status] || '?'
    puts "  #{status_icon} #{job.worker_class} [#{job.queue}] #{job.status} (#{job.attempts} attempts)"
  end

  # Recurring schedule demo
  puts "\n=== Recurring Schedule ==="
  recurring = RecurringSchedule.new(scheduler)
    .every(60,  'ReportWorker',   report_type: 'daily_summary', user_id: 0)
    .every(300, 'DataSyncWorker', source: 'prod', destination: 'backup', table: 'all')

  puts "Recurring jobs configured (not started in demo)"

  scheduler.stop(timeout: 2)
  puts "\nFinal: #{completed_count} completed, #{failed_count} failed"
end
