# ============================================================
# CLI Application Framework
# ============================================================
# A Thor/GLI-style CLI framework with subcommands, options,
# interactive prompts, progress bars, colored output,
# config files, and plugin system.
# Implements a project management CLI tool as demo.
# ============================================================

require 'optparse'
require 'json'
require 'fileutils'
require 'io/console'
require 'time'

# ── Terminal utilities ────────────────────────────────────────────────────────

module Terminal
  COLORS = {
    reset:   "\e[0m",
    bold:    "\e[1m",
    dim:     "\e[2m",
    red:     "\e[31m",
    green:   "\e[32m",
    yellow:  "\e[33m",
    blue:    "\e[34m",
    magenta: "\e[35m",
    cyan:    "\e[36m",
    white:   "\e[37m",
    gray:    "\e[90m"
  }.freeze

  def self.colorize(text, *styles)
    return text unless $stdout.tty?
    codes = styles.map { |s| COLORS[s] }.compact.join
    "#{codes}#{text}#{COLORS[:reset]}"
  end

  def self.width
    $stdout.tty? ? IO.console.winsize[1] : 80
  rescue
    80
  end

  def self.clear_line
    print "\r\e[K" if $stdout.tty?
  end

  def self.move_up(n = 1)
    print "\e[#{n}A" if $stdout.tty?
  end

  def self.hide_cursor
    print "\e[?25l" if $stdout.tty?
  end

  def self.show_cursor
    print "\e[?25h" if $stdout.tty?
  end
end

# ── Progress bar ──────────────────────────────────────────────────────────────

class ProgressBar
  def initialize(total:, label: '', width: 40, format: :bar)
    @total   = total
    @current = 0
    @label   = label
    @width   = width
    @format  = format
    @start   = Time.now
    @spinner = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
    @spin_i  = 0
  end

  def increment(n = 1)
    @current = [@current + n, @total].min
    render
  end

  def finish
    @current = @total
    render
    puts ""
  end

  def with_progress(items, &block)
    @total = items.size
    items.each_with_index do |item, i|
      block.call(item, i)
      increment
    end
    finish
  end

  private

  def render
    pct      = @total > 0 ? (@current.to_f / @total * 100).round(1) : 0
    elapsed  = Time.now - @start
    eta      = @current > 0 ? (elapsed / @current * (@total - @current)).round : 0
    filled   = (@width * @current / [@total, 1].max).round
    bar      = Terminal.colorize('█' * filled, :green) + Terminal.colorize('░' * (@width - filled), :gray)
    spinner  = @spinner[@spin_i % @spinner.size]
    @spin_i += 1

    line = "\r#{spinner} #{@label} [#{bar}] #{pct}% (#{@current}/#{@total}) ETA: #{eta}s"
    print line
    $stdout.flush
  end
end

# ── Interactive prompt ────────────────────────────────────────────────────────

module Prompt
  def self.ask(question, default: nil, required: false)
    loop do
      prompt = default ? "#{question} [#{default}]: " : "#{question}: "
      print Terminal.colorize(prompt, :cyan)
      answer = $stdin.gets&.chomp
      answer = default if answer.nil? || answer.empty?
      return answer if !required || !answer.to_s.empty?
      puts Terminal.colorize("  This field is required.", :red)
    end
  end

  def self.confirm(question, default: true)
    hint = default ? '[Y/n]' : '[y/N]'
    print Terminal.colorize("#{question} #{hint}: ", :cyan)
    answer = $stdin.gets&.chomp&.downcase
    return default if answer.nil? || answer.empty?
    %w[y yes].include?(answer)
  end

  def self.select(question, choices, default: nil)
    puts Terminal.colorize(question, :cyan)
    choices.each_with_index do |choice, i|
      marker = default == choice ? Terminal.colorize('▶', :green) : ' '
      puts "  #{marker} #{i + 1}. #{choice}"
    end
    print "Choice [1-#{choices.size}]: "
    idx = $stdin.gets&.chomp&.to_i
    idx = choices.index(default) + 1 if (idx.nil? || idx < 1 || idx > choices.size) && default
    choices[(idx || 1) - 1]
  end

  def self.password(question)
    print Terminal.colorize("#{question}: ", :cyan)
    pass = $stdin.noecho(&:gets)&.chomp
    puts ""
    pass
  end

  def self.multiline(question)
    puts Terminal.colorize("#{question} (empty line to finish):", :cyan)
    lines = []
    loop do
      line = $stdin.gets&.chomp
      break if line.nil? || line.empty?
      lines << line
    end
    lines.join("\n")
  end
end

# ── Command framework ─────────────────────────────────────────────────────────

class Command
  attr_reader :name, :description, :options, :subcommands

  def initialize(name, description: '', &block)
    @name        = name
    @description = description
    @options     = {}
    @subcommands = {}
    @action      = nil
    @before      = []
    @after       = []
    instance_eval(&block) if block
  end

  def option(flag, description: '', default: nil, required: false, type: :string)
    @options[flag] = { description: description, default: default, required: required, type: type }
  end

  def subcommand(name, description: '', &block)
    @subcommands[name.to_s] = Command.new(name, description: description, &block)
  end

  def action(&block)
    @action = block
  end

  def before(&block)
    @before << block
  end

  def after(&block)
    @after << block
  end

  def run(args, global_opts = {})
    # Check for subcommand
    if args.first && @subcommands[args.first]
      return @subcommands[args.first].run(args[1..], global_opts)
    end

    opts = parse_options(args)
    opts.merge!(global_opts)

    # Check required options
    @options.each do |flag, config|
      if config[:required] && opts[flag].nil?
        puts Terminal.colorize("Error: --#{flag} is required", :red)
        return 1
      end
    end

    @before.each { |b| b.call(opts) }
    result = @action ? @action.call(opts, args) : show_help
    @after.each  { |a| a.call(opts, result) }
    result
  end

  def show_help
    puts Terminal.colorize("Usage: #{@name}", :bold)
    puts "  #{@description}" unless @description.empty?
    unless @subcommands.empty?
      puts "\nCommands:"
      @subcommands.each { |n, cmd| puts "  #{n.ljust(20)} #{cmd.description}" }
    end
    unless @options.empty?
      puts "\nOptions:"
      @options.each do |flag, config|
        default = config[:default] ? " (default: #{config[:default]})" : ""
        req     = config[:required] ? Terminal.colorize(" [required]", :red) : ""
        puts "  --#{flag.to_s.ljust(20)} #{config[:description]}#{default}#{req}"
      end
    end
    0
  end

  private

  def parse_options(args)
    opts = @options.transform_values { |v| v[:default] }
    i = 0
    while i < args.size
      arg = args[i]
      if arg.start_with?('--')
        key = arg[2..].to_sym
        if @options[key]
          case @options[key][:type]
          when :boolean
            opts[key] = true
          else
            opts[key] = args[i + 1]
            i += 1
          end
        end
      end
      i += 1
    end
    opts
  end
end

# ── Config management ─────────────────────────────────────────────────────────

class Config
  DEFAULT_PATH = File.expand_path('~/.projectcli/config.json')

  def initialize(path = DEFAULT_PATH)
    @path = path
    @data = load
  end

  def [](key)
    @data[key.to_s]
  end

  def []=(key, value)
    @data[key.to_s] = value
    save
  end

  def get(key, default = nil)
    @data.fetch(key.to_s, default)
  end

  def set(key, value)
    self[key] = value
  end

  def delete(key)
    @data.delete(key.to_s)
    save
  end

  def all
    @data.dup
  end

  def to_s
    JSON.pretty_generate(@data)
  end

  private

  def load
    return {} unless File.exist?(@path)
    JSON.parse(File.read(@path))
  rescue
    {}
  end

  def save
    FileUtils.mkdir_p(File.dirname(@path))
    File.write(@path, JSON.pretty_generate(@data))
  end
end

# ── Project management CLI ────────────────────────────────────────────────────

class ProjectStore
  def initialize
    @projects = {}
    @tasks    = {}
    @next_id  = 1
  end

  def create_project(name:, description: '', tags: [])
    id = @next_id
    @next_id += 1
    @projects[id] = { id: id, name: name, description: description, tags: tags,
                      created_at: Time.now.iso8601, status: 'active' }
    @tasks[id] = []
    @projects[id]
  end

  def list_projects(status: nil)
    projs = @projects.values
    projs = projs.select { |p| p[:status] == status } if status
    projs
  end

  def find_project(id)
    @projects[id.to_i]
  end

  def add_task(project_id:, title:, priority: 'medium', due_date: nil)
    proj = find_project(project_id)
    return nil unless proj
    task = { id: @next_id, project_id: project_id.to_i, title: title,
             priority: priority, due_date: due_date, status: 'todo',
             created_at: Time.now.iso8601 }
    @next_id += 1
    @tasks[project_id.to_i] << task
    task
  end

  def tasks_for(project_id)
    @tasks[project_id.to_i] || []
  end

  def complete_task(task_id)
    @tasks.each_value do |tasks|
      task = tasks.find { |t| t[:id] == task_id.to_i }
      if task
        task[:status] = 'done'
        task[:completed_at] = Time.now.iso8601
        return task
      end
    end
    nil
  end

  def stats
    total_tasks = @tasks.values.flatten
    {
      projects:        @projects.size,
      active_projects: @projects.values.count { |p| p[:status] == 'active' },
      total_tasks:     total_tasks.size,
      done_tasks:      total_tasks.count { |t| t[:status] == 'done' },
      todo_tasks:      total_tasks.count { |t| t[:status] == 'todo' }
    }
  end
end

# ── Build the CLI ─────────────────────────────────────────────────────────────

def build_cli(store, config)
  root = Command.new('project', description: 'Project management CLI') do
    option :verbose,  description: 'Verbose output', type: :boolean, default: false
    option :format,   description: 'Output format (table|json|csv)', default: 'table'

    subcommand 'new', description: 'Create a new project' do
      option :name,        description: 'Project name',        required: true
      option :description, description: 'Project description', default: ''
      option :tags,        description: 'Comma-separated tags', default: ''

      action do |opts, _args|
        tags = opts[:tags].to_s.split(',').map(&:strip)
        proj = store.create_project(name: opts[:name], description: opts[:description], tags: tags)
        puts Terminal.colorize("✓ Created project ##{proj[:id]}: #{proj[:name]}", :green)
        puts "  Description: #{proj[:description]}" unless proj[:description].empty?
        puts "  Tags: #{proj[:tags].join(', ')}" unless proj[:tags].empty?
        0
      end
    end

    subcommand 'list', description: 'List all projects' do
      option :status, description: 'Filter by status (active|archived)', default: nil

      action do |opts, _args|
        projects = store.list_projects(status: opts[:status])
        if projects.empty?
          puts Terminal.colorize("No projects found.", :yellow)
        else
          puts Terminal.colorize("Projects:", :bold)
          puts Terminal.colorize("  #{'ID'.ljust(4)} #{'Name'.ljust(25)} #{'Status'.ljust(10)} Tags", :gray)
          puts Terminal.colorize("  " + "-" * 60, :gray)
          projects.each do |p|
            status_color = p[:status] == 'active' ? :green : :gray
            puts "  #{p[:id].to_s.ljust(4)} #{p[:name].ljust(25)} #{Terminal.colorize(p[:status].ljust(10), status_color)} #{p[:tags].join(', ')}"
          end
        end
        0
      end
    end

    subcommand 'task', description: 'Manage tasks' do
      subcommand 'add', description: 'Add a task to a project' do
        option :project,  description: 'Project ID', required: true
        option :title,    description: 'Task title',  required: true
        option :priority, description: 'Priority (low|medium|high)', default: 'medium'
        option :due,      description: 'Due date (YYYY-MM-DD)', default: nil

        action do |opts, _args|
          task = store.add_task(
            project_id: opts[:project],
            title:      opts[:title],
            priority:   opts[:priority],
            due_date:   opts[:due]
          )
          if task
            priority_color = { 'high' => :red, 'medium' => :yellow, 'low' => :green }[task[:priority]] || :white
            puts Terminal.colorize("✓ Task ##{task[:id]} added: #{task[:title]}", :green)
            puts "  Priority: #{Terminal.colorize(task[:priority], priority_color)}"
            puts "  Due: #{task[:due_date]}" if task[:due_date]
          else
            puts Terminal.colorize("✗ Project not found", :red)
          end
          0
        end
      end

      subcommand 'list', description: 'List tasks for a project' do
        option :project, description: 'Project ID', required: true

        action do |opts, _args|
          tasks = store.tasks_for(opts[:project])
          if tasks.empty?
            puts Terminal.colorize("No tasks for project #{opts[:project]}", :yellow)
          else
            puts Terminal.colorize("Tasks for project ##{opts[:project]}:", :bold)
            tasks.each do |t|
              icon   = t[:status] == 'done' ? Terminal.colorize('✓', :green) : Terminal.colorize('○', :gray)
              pcolor = { 'high' => :red, 'medium' => :yellow, 'low' => :green }[t[:priority]] || :white
              puts "  #{icon} ##{t[:id]} #{t[:title]} [#{Terminal.colorize(t[:priority], pcolor)}]"
            end
          end
          0
        end
      end

      subcommand 'done', description: 'Mark a task as complete' do
        option :id, description: 'Task ID', required: true

        action do |opts, _args|
          task = store.complete_task(opts[:id])
          if task
            puts Terminal.colorize("✓ Task ##{task[:id]} completed: #{task[:title]}", :green)
          else
            puts Terminal.colorize("✗ Task not found", :red)
          end
          0
        end
      end
    end

    subcommand 'stats', description: 'Show statistics' do
      action do |_opts, _args|
        s = store.stats
        puts Terminal.colorize("=== Project Statistics ===", :bold)
        puts "  Projects:    #{s[:projects]} (#{s[:active_projects]} active)"
        puts "  Tasks:       #{s[:total_tasks]} total"
        puts "  #{Terminal.colorize('✓', :green)} Done:      #{s[:done_tasks]}"
        puts "  #{Terminal.colorize('○', :gray)} Todo:      #{s[:todo_tasks]}"
        completion = s[:total_tasks] > 0 ? (s[:done_tasks].to_f / s[:total_tasks] * 100).round(1) : 0
        puts "  Completion:  #{completion}%"

        # Progress bar
        bar_width = 30
        filled = (bar_width * completion / 100).round
        bar = Terminal.colorize('█' * filled, :green) + Terminal.colorize('░' * (bar_width - filled), :gray)
        puts "  Progress:    [#{bar}]"
        0
      end
    end

    subcommand 'config', description: 'Manage configuration' do
      subcommand 'set', description: 'Set a config value' do
        option :key,   description: 'Config key',   required: true
        option :value, description: 'Config value', required: true
        action do |opts, _|
          config.set(opts[:key], opts[:value])
          puts Terminal.colorize("✓ Set #{opts[:key]} = #{opts[:value]}", :green)
          0
        end
      end

      subcommand 'get', description: 'Get a config value' do
        option :key, description: 'Config key', required: true
        action do |opts, _|
          val = config.get(opts[:key])
          puts val ? "#{opts[:key]} = #{val}" : Terminal.colorize("Key not found", :yellow)
          0
        end
      end

      subcommand 'list', description: 'List all config' do
        action do |_, _|
          puts config.to_s
          0
        end
      end
    end
  end

  root
end

# ── Demo ──────────────────────────────────────────────────────────────────────

if __FILE__ == $0
  puts Terminal.colorize("=== Project Management CLI Demo ===", :bold)
  puts ""

  store  = ProjectStore.new
  config = Config.new('/tmp/projectcli_demo.json')
  cli    = build_cli(store, config)

  # Simulate CLI commands
  commands = [
    ['new', '--name', 'Website Redesign', '--description', 'Redesign company website', '--tags', 'web,design'],
    ['new', '--name', 'API Development',  '--description', 'Build REST API',           '--tags', 'backend,api'],
    ['new', '--name', 'Mobile App',       '--description', 'iOS and Android app',      '--tags', 'mobile'],
    ['list'],
    ['task', 'add', '--project', '1', '--title', 'Create wireframes',    '--priority', 'high'],
    ['task', 'add', '--project', '1', '--title', 'Design mockups',       '--priority', 'medium'],
    ['task', 'add', '--project', '1', '--title', 'Implement frontend',   '--priority', 'high'],
    ['task', 'add', '--project', '2', '--title', 'Design API schema',    '--priority', 'high'],
    ['task', 'add', '--project', '2', '--title', 'Write endpoints',      '--priority', 'medium'],
    ['task', 'list', '--project', '1'],
    ['task', 'done', '--id', '1'],
    ['task', 'done', '--id', '3'],
    ['stats'],
    ['config', 'set', '--key', 'default_priority', '--value', 'medium'],
    ['config', 'list'],
  ]

  commands.each do |cmd|
    puts Terminal.colorize("\n$ project #{cmd.join(' ')}", :dim)
    cli.run(cmd)
  end

  # Progress bar demo
  puts "\n#{Terminal.colorize('=== Progress Bar Demo ===', :bold)}"
  bar = ProgressBar.new(total: 20, label: 'Processing files', width: 30)
  20.times do |i|
    sleep(0.05)
    bar.increment
  end
end
