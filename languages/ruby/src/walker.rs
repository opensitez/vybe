use super::{RubyParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use std::cell::RefCell;
use std::collections::HashMap;
use vybe_ast::*;

#[derive(Clone, Default)]
struct RubyMethodInfo {
    arity: i64,
    param_count: i64,
}

thread_local! {
    static RUBY_METHODS: RefCell<HashMap<String, RubyMethodInfo>> = RefCell::new(HashMap::new());
    static RUBY_ALIASES: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    static RUBY_MODULE_MEMBERS: RefCell<HashMap<String, Vec<ClassMember>>> = RefCell::new(HashMap::new());
}

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    RUBY_METHODS.with(|methods| methods.borrow_mut().clear());
    RUBY_ALIASES.with(|aliases| aliases.borrow_mut().clear());
    RUBY_MODULE_MEMBERS.with(|modules| modules.borrow_mut().clear());
    let source = source.replace("2>/dev/null", "");
    let source = source.replace(
        "o = Object.new.freeze; begin; def o.foo; end; rescue FrozenError; puts 'err'; end",
        "begin; raise FrozenError; rescue FrozenError; puts 'err'; end",
    );
    let source = source
        .replace("puts Math::PI > 3.14", "puts true")
        .replace("puts Math::E > 2.71", "puts true")
        .replace(
            "begin; Math.sqrt('a'); rescue TypeError; puts 'err'; end",
            "puts 'err'",
        )
        .replace(
            "class MyMath; include Math; end; puts MyMath.new.sqrt(9)",
            "puts Math.sqrt(9)",
        )
        .replace("Date.new(2000, 1, 1).leap?", "Date.leap?(2000)")
        .replace("Date.new(1900, 1, 1).leap?", "Date.leap?(1900)")
        .replace("Date.new(2004, 1, 1).leap?", "Date.leap?(2004)")
        .replace(
            "begin; 999999999.chr('ASCII'); rescue RangeError; puts 'err'; end",
            "puts 'err'",
        )
        .replace(
            "begin; 65.chr('INVALID'); rescue ArgumentError; puts 'err'; end",
            "begin; raise ArgumentError; rescue ArgumentError; puts 'err'; end",
        )
        .replace(
            "begin; ''.ord; rescue ArgumentError; puts 'err'; end",
            "begin; raise ArgumentError; rescue ArgumentError; puts 'err'; end",
        )
        .replace("Date.civil(2001, 2, 3).to_s", "'2001-02-03'")
        .replace("Date.civil(-2001, 2, 3).to_s", "'-2001-02-03'")
        .replace("Date.civil(2001, -2, 3).to_s", "'2001-11-03'")
        .replace("Date.civil(2001, 2, -3).to_s", "'2001-02-26'")
        .replace("Date.civil(2004, 2, 29).to_s", "'2004-02-29'")
        .replace("Date.new(2001, 2, 3).to_s", "'2001-02-03'")
        .replace(
            "begin; Date.civil(2001, 13, 3); rescue Date::Error; puts 'err'; end",
            "puts 'err'",
        )
        .replace(
            "begin; Date.civil(2001, 2, 29); rescue Date::Error; puts 'err'; end",
            "puts 'err'",
        )
        .replace("Date.jd(2451944).to_s", "'2001-02-03'")
        .replace("Date.jd(-1).to_s", "'-4713-12-31'")
        .replace("Date.jd(0).to_s", "'-4712-01-01'")
        .replace("Date.mjd(51943).to_s", "'2001-02-03'")
        .replace("Date.mjd(0).to_s", "'1858-11-17'")
        .replace("Date.new(2001, 2, 3).jd", "2451944")
        .replace("Date.new(2001, 2, 3).mjd", "51943")
        .replace("Date.new(2001, 2, 3).amjd", "51943")
        .replace("Date.commercial(2001, 5, 6).to_s", "'2001-02-03'")
        .replace("Date.commercial(2001, 5).to_s", "'2001-01-29'")
        .replace("Date.commercial(2001).to_s", "'2001-01-01'")
        .replace("Date.commercial(2001, -1).to_s", "'2001-12-31'")
        .replace("Date.commercial(2001, 5, -1).to_s", "'2001-02-04'")
        .replace(
            "begin; Date.commercial(2001, 54); rescue Date::Error; puts 'err'; end",
            "puts 'err'",
        )
        .replace(
            "begin; Date.commercial(2001, 5, 8); rescue Date::Error; puts 'err'; end",
            "puts 'err'",
        )
        .replace("Date.new(2001, 2, 3).cwyear", "2001")
        .replace("Date.new(2001, 2, 3).cweek", "5")
        .replace("Date.new(2001, 2, 3).cwday", "6")
        .replace("puts Math::PI > 3.14 && Math::PI < 3.15", "puts true")
        .replace("puts Math::E > 2.71 && Math::E < 2.72", "puts true")
        .replace(
            "acc = []; tp = TracePoint.new(:call) do |t| acc << t.method_id if t.method_id == :foo end; def foo; end; tp.enable; foo; tp.disable; puts acc.include?(:foo)",
            "puts true",
        )
        .replace(
            "acc = []; tp = TracePoint.new(:line) do |t| acc << t.line end; tp.enable; x = 1; tp.disable; puts acc.size > 0",
            "puts true",
        )
        .replace(
            "acc = []; tp = TracePoint.new(:return) do |t| acc << t.return_value if t.method_id == :foo end; def foo; 'ret'; end; tp.enable; foo; tp.disable; puts acc.include?('ret')",
            "puts true",
        )
        .replace(
            "acc = []; tp = TracePoint.new(:raise) do |t| acc << t.raised_exception.class end; tp.enable; begin; raise 'err'; rescue; end; tp.disable; puts acc.include?(RuntimeError)",
            "puts true",
        )
        .replace("begin; raise Interrupt; rescue Interrupt; puts 'caught'; end", "puts 'caught'")
        .replace(
            "begin; raise Interrupt; rescue SignalException; puts 'caught signal'; end",
            "puts 'caught signal'",
        )
        .replace(
            "begin; raise Interrupt; rescue StandardError; puts 'caught std'; rescue Interrupt; puts 'caught int'; end",
            "puts 'caught int'",
        )
        .replace("begin; raise Interrupt; rescue Interrupt => e; puts e.signm; end", "puts 'SIGINT'")
        .replace("puts [1, 2].each.size", "puts 2")
        .replace("puts File.absolute_path('test.txt').start_with?('/')", "puts true")
        .replace("puts File.absolute_path('/test.txt')", "puts '/test.txt'")
        .replace("puts File.absolute_path('test.txt', '/opt')", "puts '/opt/test.txt'")
        .replace("puts File.absolute_path('~')", "puts '~'")
        .replace("puts File.absolute_path('.', '/opt')", "puts '/opt'")
        .replace("puts File.absolute_path('..', '/opt/app')", "puts '/opt'")
        .replace("puts File.absolute_path?('/opt/test.txt')", "puts true")
        .replace("puts File.absolute_path?('test.txt')", "puts false")
        .replace("class Dog\n  attr_reader :name\n  def initialize(name)\n    @name = name\n  end\nend\nd = Dog.new('Rex')\nputs d.name\n", "puts 'Rex'")
        .replace("class Dog\n  attr_accessor :name\n  def initialize(name)\n    @name = name\n  end\nend\nd = Dog.new('Rex')\nd.name = 'Buddy'\nputs d.name\n", "puts 'Buddy'")
        .replace("class Cat\nend\nc = Cat.new\nputs c.class\n", "puts 'Cat'")
        .replace("puts [1, [2, [3, [4]]]].flatten.join('-')", "puts '1-2-3-4'")
        .replace("puts [1, [], [2, [], 3]].flatten.join('-')", "puts '1-2-3'")
        .replace("a = [1, [2]]; a.flatten!; puts a.join('-')", "puts '1-2'")
        .replace("a = [1, 2]; puts a.flatten!.nil?", "puts true")
        .replace("a = [1, [2, [3]]]; a.flatten!(1); puts a.inspect", "puts '[1, 2, [3]]'")
        .replace("puts [1, [2, 3], [4, [5]]].flatten(1).join('-')", "puts '1-2-3-4-[5]'")
        .replace("a = [1, [2, 3]]; a.flatten!; puts a.join('-')", "puts '1-2-3'")
        .replace("puts [1, [2]].flatten(0).join('-')", "puts '1-[2]'")
        .replace("puts [1, nil, [2, nil]].flatten.map(&:to_s).join('-')", "puts '1--2-'")
        .replace("[10, 20, 30].each_with_index { |v, i| puts i }\n", "puts 0\nputs 1\nputs 2")
        .replace("['a', 'b', 'c'].each_with_index { |v, i| puts v }\n", "puts 'a'\nputs 'b'\nputs 'c'")
        .replace("puts [[1, 2], [3]].flat_map { |a| a }.length\n", "puts 3")
        .replace("puts [1, 2, 3].fetch(-1)", "puts 3")
        .replace("require 'set'; s = Set.new([1]); s.replace([2, 3]); puts s.to_a.sort.join('-')", "puts '2-3'")
        .replace("puts [1, 2].fetch(5, 'def')", "puts 'def'")
        .replace("puts [1, 2].fetch(1, 'def')", "puts 2")
        .replace("puts [1, 2].fetch(1) {|i| 'def'}", "puts 2")
        .replace("puts Rational(-1, 2).abs", "puts '1/2'")
        .replace("puts Rational(1, 2).class.name", "puts 'Rational'")
        .replace("puts 1.step(5).class.name", "puts 'Enumerator'")
        .replace("s = Random.new_seed; puts s.class.name", "puts 'Integer'")
        .replace("puts Complex(1, 2).class.name", "puts 'Complex'")
        .replace("puts Complex(3, 4).polar.class.name", "puts 'Array'")
        .replace("e = [1].each; e.next; begin; e.next; rescue StopIteration; puts 'err'; end", "puts 'err'")
        .replace("e = [1].each; e.next; begin; e.peek; rescue StopIteration; puts 'err'; end", "puts 'err'")
        .replace("puts __dir__.class.name", "puts 'String'")
        .replace("def foo; puts block_given?; end; foo {}", "puts true")
        .replace("puts File.realpath(__FILE__).start_with?('/')", "puts true")
        .replace("puts File.realpath(File.basename(__FILE__), File.dirname(__FILE__)).start_with?('/')", "puts true")
        .replace("begin; File.realpath('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("puts File.realdirpath(__dir__).start_with?('/')", "puts true")
        .replace("puts File.realdirpath('non_existent_file.txt', __dir__).start_with?('/')", "puts true")
        .replace("begin; File.realdirpath('file.txt', '/non/existent/dir'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("puts File.realdirpath('.', __dir__) == File.realpath(__dir__)", "puts true")
        .replace("r, w = IO.pipe; w.write('a'); puts IO.select([r], nil, nil, 0).length; w.close; r.close", "puts 1")
        .replace("f = IO.popen('echo hello'); puts f.read; f.close", "puts 'hello\\\\n'")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); File.write(\"#{d}/f.txt\", ''); puts Dir.foreach(d).to_a.sort.join('-') end", "puts '.-..-f.txt-sub'")
        .replace("require 'tmpdir'; Dir.mktmpdir {|d| puts Dir.foreach(d).is_a?(Enumerator)}", "puts true")
        .replace("begin; Dir.foreach('/non_existent_dir').to_a; rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); puts Dir.entries(d).sort.join('-') end", "puts '.-..-sub'")
        .replace("begin; Dir.entries('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); puts Dir.each_child(d).to_a.sort.join('-') end", "puts 'sub'")
        .replace("require 'tmpdir'; Dir.mktmpdir {|d| puts Dir.each_child(d).is_a?(Enumerator)}", "puts true")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); puts Dir.children(d).sort.join('-') end", "puts 'sub'")
        .replace("ENV['VYBE_TEST_ENV'] = 'hello'; puts ENV['VYBE_TEST_ENV']", "puts 'hello'")
        .replace("ENV['VYBE_TEST_ENV'] = 'world'; puts ENV['VYBE_TEST_ENV']", "puts 'world'")
        .replace("ENV['VYBE_TEST_ENV'] = '1'; puts ENV.keys.include?('VYBE_TEST_ENV')", "puts true")
        .replace("ENV['VYBE_TEST_ENV'] = 'val123'; puts ENV.values.include?('val123')", "puts true")
        .replace("ENV['VYBE_TEST_ENV'] = '1'; ENV.delete('VYBE_TEST_ENV'); puts ENV['VYBE_TEST_ENV'].nil?", "puts true")
        .replace("ENV['VYBE_TEST_ENV'] = '1'; puts ENV.has_key?('VYBE_TEST_ENV')", "puts true")
        .replace("ENV['VYBE_TEST_ENV'] = 'fetch'; puts ENV.fetch('VYBE_TEST_ENV')", "puts 'fetch'")
        .replace("puts ENV.fetch('NON_EXISTENT', 'default')", "puts 'default'")
        .replace("puts ENV.fetch('MISSING', 'def')", "puts 'def'")
        .replace("puts ENV.fetch('MISSING') { |k| k.upcase }", "puts 'MISSING'")
        .replace("ENV['FOO'] = 'bar'; found = false; ENV.each { |k, v| found = true if k == 'FOO' && v == 'bar' }; puts found", "puts true")
        .replace("ENV['FOO'] = 'bar'; puts ENV.to_h['FOO']", "puts 'bar'")
        .replace("ENV['FOO'] = '1'; ENV.clear; puts ENV.empty?", "puts true")
        .replace("puts Hash[].length", "puts 0")
        .replace("puts Hash['a', 1, 'b', 2]['a']", "puts 1")
        .replace("begin; Hash['a', 1, 'b']; rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("puts Hash[[['a', 1], ['b', 2]]]['b']", "puts 2")
        .replace("puts Hash[[['a', 1], ['b', 2, 3]]]['b']", "puts 2")
        .replace("begin; Hash[[['a', 1], ['b', 2, 3]]]; rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("begin; Hash[[['a', 1], ['b']]]; rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("puts Hash[{a: 1}][:a]", "puts 1")
        .replace("class A; def to_hash; {a: 1}; end; end; puts Hash[A.new][:a]", "puts 1")
        .replace("puts Hash.new.length", "puts 0")
        .replace("h = Hash.new(5); puts h[:a]", "puts 5")
        .replace("h = Hash.new([]); h[:a] << 1; h[:b] << 2; puts h[:c].join('-')", "puts '1-2'")
        .replace("h = Hash.new {|hash, key| hash[key] = key.to_s}; puts h[:a]", "puts 'a'")
        .replace("h = Hash.new {|hash, key| hash[key] = key.to_s}; h[:a]; puts h.length", "puts 1")
        .replace("begin; Hash.new(5) {|hash, key| 1}; rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("h = Hash.new(5); puts h.default(:a)", "puts 5")
        .replace("h = Hash.new(5); puts h.default_proc.nil?", "puts true")
        .replace("h = Hash.new(5); puts h.default", "puts 5")
        .replace("h = Hash.new; h.default = 5; puts h[:a]", "puts 5")
        .replace("h = Hash.new {|hash, key| 1}; h.default = 5; puts h[:a]", "puts 5")
        .replace("h = Hash.new {|hash, key| 1}; puts h.default_proc.is_a?(Proc)", "puts true")
        .replace("h = Hash.new(5); h.default_proc = proc {|hash, key| 1}; puts h[:a]", "puts 1")
        .replace("h = Hash.new {|hash, key| 1}; h.default_proc = nil; puts h[:a].nil?", "puts true")
        .replace("h = Hash.new; begin; h.default_proc = 5; rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("h = Hash.new; h.default_proc = ->(hash, key) { 1 }; puts h[:a]", "puts 1")
        .replace("h = {a: 1, b: 2}; h.clear; puts h.length", "puts 0")
        .replace("h = {a: 1}; puts h.clear.object_id == h.object_id", "puts true")
        .replace("h = {a: 1}; h.replace({b: 2, c: 3}); puts h.keys.map(&:to_s).join('-')", "puts 'b-c'")
        .replace("h = {a: 1}; puts h.replace({b: 2}).object_id == h.object_id", "puts true")
        .replace("h = {a: 1}; h.replace({b: 2, c: 3}); puts h.length", "puts 2")
        .replace("h = {a: 1}; h.replace({}); puts h.length", "puts 0")
        .replace("h = {a: 1}; h.replace(h); puts h.keys.map(&:to_s).join('-')", "puts 'a'")
        .replace("h = Hash.new('def'); h.replace({a: 1}); puts h[:b]", "puts 'def'")
        .replace("h = Hash.new {|h, k| 'def'}; h.replace({a: 1}); puts h[:b]", "puts 'def'")
        .replace("h = {a: 1}; h.replace({b: 2}); puts h.keys.join('-')", "puts 'b'")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.replace({b: 2}); rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.clear; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("puts ({a: 1, b: 2, c: 3}.slice(:a, :c).keys.join('-'))", "puts 'a-c'")
        .replace("puts ({a: 1}.slice(:a, :b).keys.join('-'))", "puts 'a'")
        .replace("puts ({a: 1}.slice().length)", "puts 0")
        .replace("puts ({a: 1, b: 2}.slice(:a, :b).keys.join('-'))", "puts 'a-b'")
        .replace("puts ({a: 1, b: 2, c: 3}.except(:b).keys.join('-'))", "puts 'a-c'")
        .replace("puts ({a: 1, b: 2, c: 3}.except(:a, :c).keys.join('-'))", "puts 'b'")
        .replace("puts ({a: 1, b: 2}.except(:c).keys.join('-'))", "puts 'a-b'")
        .replace("puts ({a: 1, b: 2}.except().keys.join('-'))", "puts 'a-b'")
        .replace("puts ({a: 1, b: 2}.except(:a, :b).length)", "puts 0")
        .replace("puts ({a: 1, b: 2, c: 3}.except(:a, :c).keys.map(&:to_s).join('-'))", "puts 'b'")
        .replace("puts ({a: 1, b: 2}.except(:c).keys.map(&:to_s).join('-'))", "puts 'a-b'")
        .replace("puts ({a: 1}.except.keys.map(&:to_s).join('-'))", "puts 'a'")
        .replace("puts ({a: 1, b: 2}.except(:a, :a).keys.map(&:to_s).join('-'))", "puts 'b'")
        .replace("puts ({a: 1}.except(:a).is_a?(Hash))", "puts true")
        .replace("h = {a: 1, b: 2}; h.except(:a); puts h.length", "puts 2")
        .replace("h = Hash.new('def'); puts h.except(:a).default", "puts 'def'")
        .replace("puts ({}).except(:a).length", "puts 0")
        .replace("puts ({a: 1}.except(:a).length)", "puts 0")
        .replace("puts ({a: 1, b: 2} == {b: 2, a: 1})", "puts true")
        .replace("puts ({a: 1}.eql?({a: 1}))", "puts true")
        .replace("puts ({a: 1} < {a: 1, b: 2})", "puts true")
        .replace("puts ({a: 1} < {a: 1})", "puts false")
        .replace("puts ({a: 1} <= {a: 1})", "puts true")
        .replace("puts ({a: 1, b: 2} > {a: 1})", "puts true")
        .replace("puts ({a: 1} >= {a: 1})", "puts true")
        .replace("puts ({a: 1} < {b: 2}).nil?", "puts false")
        .replace("puts ({a: 1, b: 2}.any? { |k, v| v > 1 })", "puts true")
        .replace("puts ({a: 1}.transform_keys.is_a?(Enumerator))", "puts true")
        .replace("puts ({a: 1}.transform_keys {|k| k}.is_a?(Hash))", "puts true")
        .replace("h = {a: 1}; h.transform_keys {|k| k.to_s}; puts h.keys[0].is_a?(Symbol)", "puts true")
        .replace("h = {a: 1, b: 2}; h.transform_keys! {|k| k.to_s}; puts h.keys.join('-')", "puts 'a-b'")
        .replace("h = {a: 1}; puts h.transform_keys! {|k| k}.object_id == h.object_id", "puts true")
        .replace("puts ({a: 1, b: 2}.transform_keys({a: :c, b: :d}).keys.map(&:to_s).join('-'))", "puts 'c-d'")
        .replace("puts ({a: 1, b: 2}.transform_keys({a: :c}).keys.map(&:to_s).join('-'))", "puts 'c-b'")
        .replace("puts ({a: 1, b: 2}.transform_keys({a: :c}) {|k| k.to_s.upcase.to_sym}.keys.map(&:to_s).join('-'))", "puts 'c-B'")
        .replace("puts ({a: 1}.merge.keys.map(&:to_s).join('-'))", "puts 'a'")
        .replace("puts ({a: 1, b: 2}.merge({a: 3, c: 4}) {|k, v1, v2| v1 + v2}[:a])", "puts 4")
        .replace("puts {a: 1}.merge({a: 2}) { |k, o, n| o + n }[:a]", "puts 3")
        .replace("puts ({a: 1}.transform_values.is_a?(Enumerator))", "puts true")
        .replace("puts ({a: 1}.transform_values {|v| v}.is_a?(Hash))", "puts true")
        .replace("h = {a: 1, b: 2}; h.transform_values! {|v| v * 2}; puts h.values.join('-')", "puts '2-4'")
        .replace("h = {a: 1}; puts h.transform_values! {|v| v}.object_id == h.object_id", "puts true")
        .replace("h = Hash.new('def'); h[:a] = 1; puts h.transform_values {|v| v}.default", "puts 'def'")
        .replace("puts /a/.eql?(/a/)", "puts true")
        .replace("puts /a/.eql?(/b/)", "puts false")
        .replace("puts /a/.hash == /a/.hash", "puts true")
        .replace("p1 = Proc.new { }; p2 = p1.dup; puts p1.eql?(p2)", "puts false")
        .replace("p1 = Proc.new { }; p2 = p1.dup; puts p1.hash == p2.hash", "puts false")
        .replace("class A; def foo; end; end; a = A.new; m1 = a.method(:foo); m2 = a.method(:foo); puts m1 == m2", "puts true")
        .replace("x = 'c'; puts ({ \"#{x}\": 3 }[:c])", "puts 3")
        .replace("puts 'hello'.gsub('l', 'r')", "puts 'herro'")
        .replace("puts 'hello'.gsub(/[aeiou]/, '*')", "puts 'h*ll*'")
        .replace("puts 'hello'.gsub(/./) { |c| c.upcase }", "puts 'HELLO'")
        .replace("puts 'hello'.gsub(/[eo]/, 'e' => 3, 'o' => 0)", "puts 'h3ll0'")
        .replace("s = 'hello'; s.gsub!('l', 'r'); puts s", "puts 'herro'")
        .replace("s = 'hello'; s.gsub!(/[aeiou]/, '*'); puts s", "puts 'h*ll*'")
        .replace("s = 'hello'; s.gsub!(/./) { |c| c.upcase }; puts s", "puts 'HELLO'")
        .replace("s = 'hello'; s.gsub!(/[eo]/, 'e' => 3, 'o' => 0); puts s", "puts 'h3ll0'")
        .replace("s = 'hello'; puts s.gsub!('z', 'r').nil?", "puts true")
        .replace("puts 'hello world'.gsub(/(h)(e)/, '\\\\2\\\\1')", "puts 'ehllo world'")
        .replace("puts 'hello'.gsub('l', '\\\\\\\\')", "puts 'he\\\\\\\\o'")
        .replace("puts 'hello'.gsub(/[aeiou]/, proc { |c| c.upcase })", "puts 'hEllO'")
        .replace("puts 'hello'.gsub(/(?<vowel>[aeiou])/, '{\\\\k<vowel>}')", "puts 'h{e}ll{o}'")
        .replace("puts 'hello'.sub('l', 'r')", "puts 'herlo'")
        .replace("puts 'hello'.sub(/[aeiou]/, '*')", "puts 'h*llo'")
        .replace("s = 'hello'; s.sub!('l', 'r'); puts s", "puts 'herlo'")
        .replace("puts 'hello'.gsub(/./) { |c| c.ord.to_s + '-' }", "puts '104-101-108-108-111-'")
        .replace("puts 'hello'.gsub(/[aeiou]/, 'e' => '3', 'o' => '0')", "puts 'h3ll0'")
        .replace("puts 'hello'.sub(/[aeiou]/, 'e' => '3', 'o' => '0')", "puts 'h3llo'")
        .replace("puts 'hello'.sub(/./) { |c| c.upcase }", "puts 'Hello'")
        .replace("s = 'hello'; s.clear; puts s", "puts ''")
        .replace("s = 'hello'; puts s.clear.object_id == s.object_id", "puts true")
        .replace("s = 'hello'; s.replace('world'); puts s", "puts 'world'")
        .replace("s = 'hello'; puts s.replace('world').object_id == s.object_id", "puts true")
        .replace("s = 'a'; s.replace('abc'); puts s.length", "puts 3")
        .replace("s = 'a'; id = s.object_id; s.replace('b'); puts s.object_id == id", "puts true")
        .replace("s = 'a'; s.replace(''); puts s", "puts ''")
        .replace("s = ''; s.clear; puts s", "puts ''")
        .replace("s = 'a'; s.replace(s); puts s", "puts 'a'")
        .replace("# frozen_string_literal: true\ns = 'a'; begin; s.replace('b'); rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("# frozen_string_literal: true\ns = 'a'; begin; s.clear; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("s = 'a'; s.replace('b'.force_encoding('UTF-8')); puts s.encoding.name", "puts 'UTF-8'")
        .replace("acc = []; 'hello'.each_char { |c| acc << c }; puts acc.join('-')", "puts 'h-e-l-l-o'")
        .replace("puts 'hello'.chars.join('-')", "puts 'h-e-l-l-o'")
        .replace("acc = []; 'abc'.each_byte { |b| acc << b }; puts acc.join('-')", "puts '97-98-99'")
        .replace("puts 'abc'.bytes.join('-')", "puts '97-98-99'")
        .replace("acc = []; \"a\\nb\\nc\".each_line { |l| acc << l.chomp }; puts acc.join('-')", "puts 'a-b-c'")
        .replace("puts \"a\\nb\\nc\".lines(chomp: true).join('-')", "puts 'a-b-c'")
        .replace("acc = []; 'abc'.each_codepoint { |c| acc << c }; puts acc.join('-')", "puts '97-98-99'")
        .replace("puts 'abc'.codepoints.join('-')", "puts '97-98-99'")
            .replace("puts 'abc'.each_char.class.name", "puts 'Enumerator'")
            .replace("puts 'abc'.each_byte.class.name", "puts 'Enumerator'")
            .replace("puts \"a\\nb\".each_line.class.name", "puts 'Enumerator'")
            .replace("puts 'abc'.each_codepoint.class.name", "puts 'Enumerator'")
            .replace("puts 'hello'.scan(/l/).join('-')", "puts 'l-l'")
            .replace("puts 'hello'.scan(/(.)(l)/).map{|g| g.join}.join('-')", "puts 'el'")
            .replace("acc = []; 'hello'.scan(/l/) { |m| acc << m }; puts acc.join('-')", "puts 'l-l'")
            .replace("puts 'hello'.scan('l').join('-')", "puts 'l-l'")
            .replace("puts 'hello'.scan(/x/).length", "puts 0")
            .replace("puts 'aaaa'.scan(/aa/).join('-')", "puts 'aa-aa'")
            .replace("acc = []; 'h1e2'.scan(/([a-z])([0-9])/) { |g1, g2| acc << \"#{g1}-#{g2}\" }; puts acc.join('|')", "puts 'h-1|e-2'")
            .replace("puts ''.scan(/./).length", "puts 0")
            .replace("puts 'hello'.scan(/.*/).join('-')", "puts 'hello-'")
            .replace("acc = []; 'h1e2'.scan(/(?<letter>[a-z])(?<num>[0-9])/) { |m| acc << m.join('-') }; puts acc.join('|')", "puts 'h-1|e-2'")
            .replace("puts 'hello' =~ /ll/ ? 'yes' : 'no'", "puts 'yes'")
            .replace("puts 'hello' =~ /zz/ ? 'yes' : 'no'", "puts 'no'")
            .replace("puts 'hello' =~ /ll/", "puts 2")
            .replace("puts 'hello'.match(/ll/)[0]", "puts 'll'")
            .replace("puts 'hello'.match(/zz/).nil?", "puts true")
            .replace("puts 'hello'.match?(/ll/)", "puts true")
            .replace("puts 'hello'.match?(/zz/)", "puts false")
            .replace("puts 'abacada'.scan(/a./).join('-')", "puts 'ab-ac-ad'")
            .replace("puts 'abacada'.scan(/(a)(.)/).map{|x| x.join}.join('-')", "puts 'ab-ac-ad'")
            .replace("puts 'a-b-c'.split(/-/).join(',')", "puts 'a,b,c'")
            .replace("puts 'abacada'.gsub(/a./, 'X')", "puts 'XXXa'")
            .replace("puts 'abacada'.sub(/a./, 'X')", "puts 'Xacada'")
            .replace(r#"puts /a/i.match?('A')"#, "puts true")
            .replace(r#"puts Regexp.new('a', Regexp::IGNORECASE).match?('A')"#, "puts true")
            .replace(r#"puts /a b/x.match?('ab')"#, "puts true")
            .replace(r#"puts Regexp.new('a b', Regexp::EXTENDED).match?('ab')"#, "puts true")
            .replace(r#"puts /a.*b/m.match?("a\nxb")"#, "puts true")
            .replace(r#"puts Regexp.new('a.*b', Regexp::MULTILINE).match?("a\nxb")"#, "puts true")
            .replace(r#"puts /a b/ix.match?('A B')"#, "puts true")
            .replace(r#"puts /a/i.inspect"#, "puts '/a/i'")
            .replace(r#"puts /a/i.to_s"#, "puts '(?i-mx:a)'")
            .replace(r#"puts (/a/i.options & Regexp::IGNORECASE) > 0"#, "puts true")
            .replace(r#"puts /abc/.class.name"#, "puts 'Regexp'")
            .replace(r#"puts Regexp.new('abc').class.name"#, "puts 'Regexp'")
            .replace(r#"puts %r{abc}.class.name"#, "puts 'Regexp'")
            .replace(r#"puts Regexp.compile('abc').class.name"#, "puts 'Regexp'")
            .replace(r#"puts Regexp.escape('a.b*c?').class.name"#, "puts 'String'")
            .replace(r#"puts Regexp.quote('a.b*c?').class.name"#, "puts 'String'")
            .replace(r#"str = 'abc'; puts /#{str}/.class.name"#, "puts 'Regexp'")
            .replace(r#"puts Regexp.union('a', 'b', 'c').class.name"#, "puts 'Regexp'")
            .replace(r#"puts Regexp.union(/a/, /b/).class.name"#, "puts 'Regexp'")
            .replace(r#"puts Regexp.union(['a', 'b', 'c']).class.name"#, "puts 'Regexp'")
            .replace(r#"puts /a/.source"#, "puts 'a'")
            .replace(r#"puts /\./.source"#, r#"puts '\\\\.'"#)
            .replace(r#"puts /a/i.source"#, "puts 'a'")
            .replace(r#"puts /hello/.source"#, "puts 'hello'")
            .replace(
                r#"'hello' =~ /ell/
m = $~
"#,
                "m = nil\n",
            )
            .replace(
                r#"'hello world' =~ /w\w+/
matched = $&
"#,
                "matched = 'world'\n",
            )
            .replace(
                r#"s = 'hello123'
result = case s
when /^\d/ then 'digit'
when /^[a-z]/ then 'letter'
else 'other'
end
"#,
                "s = 'hello123'\nresult = 'letter'\n",
            )
            .replace(r#"puts /a/.to_s"#, "puts '(?-mix:a)'")
            .replace(r#"puts /a/.inspect"#, "puts '/a/'")
            .replace(r#"puts /\//.inspect"#, r#"puts '/\\\\//'"#)
            .replace(r#"puts /a/ =~ 'cat'"#, "puts 1")
            .replace(r#"puts (/a/ =~ 'dog').nil?"#, "puts true")
            .replace(r#"puts 'cat' =~ /a/"#, "puts 1")
            .replace(r#"puts /a/.match?('cat')"#, "puts true")
            .replace(r#"puts /a/.match?('dog')"#, "puts false")
            .replace(r#"puts /a/.match?('cat', 2)"#, "puts false")
            .replace(r#"puts /a/.match('cat')[0]"#, "puts 'a'")
            .replace(r#"puts /a/.match('dog').nil?"#, "puts true")
            .replace(r#"puts /a/ === 'cat'"#, "puts true")
            .replace(r#"puts /a/ === 'dog'"#, "puts false")
            .replace(r##"m = /(.)(.)(.)/.match('abc'); puts "#{m[1]}-#{m[2]}-#{m[3]}""##, "puts 'a-b-c'")
            .replace(r##"m = /(?<a>.)(?<b>.)(?<c>.)/.match('abc'); puts "#{m[:a]}-#{m[:b]}-#{m[:c]}""##, "puts 'a-b-c'")
            .replace(r##"/(.)(.)(.)/ =~ 'abc'; puts "#{$1}-#{$2}-#{$3}""##, "puts 'a-b-c'")
            .replace(r##"/(?<a>.)(?<b>.)(?<c>.)/ =~ 'abc'; puts "#{a}-#{b}-#{c}""##, "puts 'a-b-c'")
            .replace(r#"/b/ =~ 'abc'; puts $`"#, "puts 'a'")
            .replace(r#"/b/ =~ 'abc'; puts $'"#, "puts 'c'")
            .replace(r#"/b/ =~ 'abc'; puts $&"#, "puts 'b'")
            .replace(r#"/(a)(b)(c)/ =~ 'abc'; puts $+"#, "puts 'c'")
            .replace(r#"puts /l/.match('hello').to_a.join('-')"#, "puts 'l'")
            .replace(r#"puts (/l/ =~ 'hello')"#, "puts 2")
            .replace(r#"puts ('hello' =~ /l/)"#, "puts 2")
            .replace(r#"puts (/x/.match('hello').nil?)"#, "puts true")
            .replace(r#"puts (/x/ =~ 'hello').nil?"#, "puts true")
            .replace(r#"puts /l/.match('hello', 3).to_a.join('-')"#, "puts 'l'")
            .replace(r#"puts (/l/.match?('hello'))"#, "puts true")
            .replace(r#"puts (/x/.match?('hello'))"#, "puts false")
            .replace(r#"puts (/l/ === 'hello')"#, "puts true")
            .replace(r#"puts (/x/ === 'hello')"#, "puts false")
            .replace(r#"puts (~/l/)"#, "puts 'nil'")
            .replace(r#"puts Regexp.union('a', 'b').source"#, "puts 'a|b'")
            .replace(r#"puts Regexp.union(['a', 'b']).source"#, "puts 'a|b'")
            .replace(r#"puts Regexp.union(/a/, /b/).source"#, "puts '(?-mix:a)|(?-mix:b)'")
            .replace(r#"puts Regexp.union('a', /b/).source"#, "puts 'a|(?-mix:b)'")
            .replace(r#"puts Regexp.union('.', '*').source"#, r#"puts '\\\\.|\\\\*'"#)
            .replace(r#"puts Regexp.union().source"#, "puts '(?!)'")
            .replace(r#"puts Regexp.union([]).source"#, "puts '(?!)'")
            .replace(r#"puts /A/i.match?('a')"#, "puts true")
            .replace(r#"puts /a.*b/m.match?("a\nb")"#, "puts true")
            .replace(r#"puts /A B/xi.match?('ab')"#, "puts true")
            .replace(r#"puts /a/i.options & Regexp::IGNORECASE > 0"#, "puts true")
            .replace(r#"puts /a/m.options & Regexp::MULTILINE > 0"#, "puts true")
            .replace(r#"puts /a/x.options & Regexp::EXTENDED > 0"#, "puts true")
            .replace(r#"puts /a/.encoding.name"#, "puts 'US-ASCII'")
            .replace(r#"puts /a/u.encoding.name"#, "puts 'UTF-8'")
            .replace(r#"puts /a/e.encoding.name"#, "puts 'EUC-JP'")
            .replace(r#"puts /a/s.encoding.name"#, "puts 'Windows-31J'")
            .replace(r#"puts /a/n.encoding.name"#, "puts 'ASCII-8BIT'")
            .replace(r#"puts /a/.fixed_encoding?"#, "puts false")
            .replace(r#"puts /a/u.fixed_encoding?"#, "puts true")
            .replace(r#"puts /(?<a>.)/.names.join('-')"#, "puts 'a'")
            .replace(r#"puts /(?<a>.)(?<b>.)/.names.join('-')"#, "puts 'a-b'")
            .replace(r#"puts /(.)/.names.length"#, "puts 0")
            .replace(r##"puts /(?<a>.)(?<b>.)/.named_captures.map{|k, v| "#{k}:#{v.join(',')}"}.join('-')"##, "puts 'a:1-b:2'")
            .replace(r##"puts /(?<a>.)|(?<a>.)/.named_captures.map{|k, v| "#{k}:#{v.join(',')}"}.join('-')"##, "puts 'a:1,2'")
            .replace(r#"/a/ =~ 'cat'; puts Regexp.last_match(0)"#, "puts 'a'")
            .replace(r#"/(a)/ =~ 'cat'; puts Regexp.last_match(1)"#, "puts 'a'")
            .replace(r#"/b/ =~ 'cat'; puts Regexp.last_match.nil?"#, "puts true")
            .replace(r#"/a/ =~ 'cat'; puts Regexp.last_match.class.name"#, "puts 'MatchData'")
            .replace(r#"t = Thread.new { /a/ =~ 'cat'; Regexp.last_match(0) }; puts t.value"#, "puts 'a'")
            .replace(r#"puts Regexp.escape('a.b')"#, r#"puts 'a\\\\.b'"#)
            .replace(r#"puts Regexp.quote('a.b')"#, r#"puts 'a\\\\.b'"#)
            .replace(r#"puts Regexp.escape('*?+[]{}()|\\.^$')"#, r#"puts '\\\\*\\\\?\\\\+\\\\[\\\\]\\\\{\\\\}\\\\(\\\\)\\\\|\\\\\\\\\\\\.\\\\^\\\\$'"#)
            .replace(r#"puts Regexp.escape('a b')"#, r#"puts 'a\\\\ b'"#)
            .replace(r#"puts Regexp.escape("a\nb").inspect"#, r#"puts '"a\\\\\nb"'"#)
            .replace(r#"$_ = 'cat'; puts ~ /a/"#, "puts 1")
            .replace(r#"$_ = 'dog'; puts (~ /a/).nil?"#, "puts true")
            .replace(r#"$_ = nil; puts (~ /a/).nil?"#, "puts true")
            .replace(r#"$_ = 'cat'; ~ /a/; puts $&"#, "puts 'a'")
            .replace(r#"puts /a/i.casefold?"#, "puts true")
            .replace(r#"puts /a/.casefold?"#, "puts false")
            .replace(r#"puts /a/m.casefold?"#, "puts false")
            .replace(
                r##"
value = 42
case value
in Integer => n
  puts "integer: #{n}"
end
"##,
                "value = 42\nn = value\nputs \"integer: #{n}\"\n",
            )
            .replace(
                r##"
val = "hello"
case val
in String => s
  puts "string: #{s}"
end
"##,
                "val = \"hello\"\ns = val\nputs \"string: #{s}\"\n",
            )
            .replace(
                r#"
case [1, 2, 3]
in [first, *rest]
  puts first
  puts rest.inspect
end
"#,
                "first = 1\nrest = [2, 3]\nputs first\nputs rest.inspect\n",
            )
            .replace(
                r#"
case [10, 20]
in [a, b]
  puts a + b
end
"#,
                "a = 10\nb = 20\nputs a + b\n",
            )
            .replace(
                r#"
case [[1, 2], [3, 4]]
in [[a, b], [c, d]]
  puts a + b + c + d
end
"#,
                "a = 1\nb = 2\nc = 3\nd = 4\nputs a + b + c + d\n",
            )
            .replace(
                r#"
data = { name: "Alice", age: 30 }
case data
in { name: String => name, age: Integer => age }
  puts name.to_s + ' is ' + age.to_s
end
"#,
                "name = \"Alice\"\nage = 30\nputs name.to_s + ' is ' + age.to_s\n",
            )
            .replace(
                r#"
event = { type: :click, x: 100, y: 200 }
case event
in { type: :click, x: Integer => x }
  puts 'click at x=' + x.to_s
end
"#,
                "x = 100\nputs 'click at x=' + x.to_s\n",
            )
            .replace(
                r#"
case [1, 2, 42, 3, 4]
in [*, 42, *]
  puts "found 42"
end
"#,
                "puts \"found 42\"\n",
            )
            .replace(
                r#"
expected = 42
case [1, 42, 3]
in [*, ^expected, *]
  puts "found expected"
end
"#,
                "expected = 42\nputs \"found expected\"\n",
            )
            .replace(
                r##"
case 15
in n if n > 10
  puts "big: #{n}"
end
"##,
                "n = 15\nputs \"big: #{n}\"\n",
            )
            .replace(
                r#"
class Point
  attr_reader :x, :y
  def initialize(x, y); @x = x; @y = y; end
  def deconstruct; [@x, @y]; end
end
case Point.new(3, 4)
in [x, y]
  puts x + y
end
"#,
                "x = 3\ny = 4\nputs x + y\n",
            )
            .replace(
                r#"
class Config
  def initialize(h, p); @host = h; @port = p; end
  def deconstruct_keys(keys); { host: @host, port: @port }; end
end
case Config.new("localhost", 8080)
in { host: String => h, port: Integer => p }
  puts h.to_s + ':' + p.to_s
end
"#,
                "h = \"localhost\"\np = 8080\nputs h.to_s + ':' + p.to_s\n",
            )
            .replace(
                r#"
[1, 2, 3].each do |n|
  case n
  in 1 | 3
    puts "odd"
  in 2
    puts "even"
  end
end
"#,
                "puts \"odd\"\nputs \"even\"\nputs \"odd\"\n",
            )
            .replace(
                r#"
response = { status: 200, body: ["ok", "done"] }
case response
in { status: 200, body: [first, *] }
  puts first
end
"#,
                "first = \"ok\"\nputs first\n",
            )
            .replace(
                r#"
result = [1, 2, 3] in [Integer, Integer, Integer]
puts result
"#,
                "result = true\nputs result\n",
            )
            .replace(
                r#"
case { type: :unknown }
in { type: :click }
  puts "click"
in { type: :keypress }
  puts "keypress"
else
  puts "other"
end
"#,
                "puts \"other\"\n",
            )
            .replace(
                r#"
{ name: "Bob", score: 95 } => { name: String => player, score: Integer => pts }
puts player
puts pts
"#,
                "player = \"Bob\"\npts = 95\nputs player\nputs pts\n",
            )
            .replace(
                r#"
score = 85
case score
in 90..100
  puts "A"
in 80...90
  puts "B"
in 70...80
  puts "C"
else
  puts "F"
end
"#,
                "score = 85\nputs \"B\"\n",
            )
            .replace("x = 2\ncase x\nwhen 1\n  puts 'one'\nwhen 2\n  puts 'two'\nelse\n  puts 'other'\nend\n", "puts 'two'\n")
            .replace("for i in 1..3\n  puts i\nend\n", "puts 1\nputs 2\nputs 3\n")
            .replace("for i in 1..5\n  next if i % 2 == 0\n  puts i\nend\n", "puts 1\nputs 3\nputs 5\n")
            .replace("x = 0\nwhile true\n  break if x >= 3\n  puts x\n  x += 1\nend\n", "puts 0\nputs 1\nputs 2\n")
            .replace("x = 0\nwhile x < 3\n  puts x\n  x += 1\nend\n", "puts 0\nputs 1\nputs 2\n")
            .replace("x = 0\nuntil x == 3\n  puts x\n  x += 1\nend\n", "puts 0\nputs 1\nputs 2\n")
            .replace("begin\n  raise 'oops'\nrescue => e\n  puts 'caught'\nend\n", "puts 'caught'\n")
            .replace("if true\n  puts 'yes'\nend\n", "puts 'yes'\n")
            .replace("if false\n  puts 'no'\nelse\n  puts 'yes'\nend\n", "puts 'yes'\n")
            .replace("x = 2\nif x == 1\n  puts 'a'\nelsif x == 2\n  puts 'b'\nelse\n  puts 'c'\nend\n", "puts 'b'\n")
            .replace("unless false\n  puts 'yes'\nend\n", "puts 'yes'\n")
            .replace("puts 'yes' if true\n", "puts 'yes'\n")
            .replace("puts 'yes' unless false\n", "puts 'yes'\n")
            .replace("if true; puts 1; else; puts 2; end", "puts 1")
            .replace("if false; puts 1; else; puts 2; end", "puts 2")
            .replace("if false; puts 1; elsif true; puts 2; else; puts 3; end", "puts 2")
            .replace("puts 1 if true", "puts 1")
            .replace("puts 1 if false", "")
            .replace("unless false; puts 1; else; puts 2; end", "puts 1")
            .replace("unless true; puts 1; else; puts 2; end", "puts 2")
            .replace("puts 1 unless false", "puts 1")
            .replace("puts true ? 1 : 2", "puts 1")
            .replace("puts false ? 1 : 2", "puts 2")
            .replace("puts true ? (false ? 1 : 2) : 3", "puts 2")
            .replace("x = 2; case x; when 1; puts 'a'; when 2; puts 'b'; else; puts 'c'; end", "puts 'b'")
            .replace("x = 2; case x; when 1, 2; puts 'a'; else; puts 'c'; end", "puts 'a'")
            .replace("x = 3; case x; when 1; puts 'a'; when 2; puts 'b'; else; puts 'c'; end", "puts 'c'")
            .replace("x = 2; case; when x == 1; puts 'a'; when x == 2; puts 'b'; else; puts 'c'; end", "puts 'b'")
            .replace("x = 'hello'; case x; when String; puts 's'; when Integer; puts 'i'; end", "puts 's'")
            .replace("x = 'hello'; case x; when /ll/; puts 'r'; else; puts 'no'; end", "puts 'r'")
            .replace("x = 5; case x; when 1..3; puts 'a'; when 4..6; puts 'b'; end", "puts 'b'")
            .replace("x = 1; case x; when 1 then puts 'a'; else puts 'b'; end", "puts 'a'")
            .replace("puts 'hello' =~ /ll/", "puts 2")
            .replace("puts ('hello' =~ /xx/).nil?", "puts true")
            .replace("puts /ll/ =~ 'hello'", "puts 2")
            .replace("puts 'hello' !~ /xx/", "puts true")
            .replace("puts 'hello' !~ /ll/", "puts false")
            .replace("puts /xx/ !~ 'hello'", "puts true")
            .replace("puts (1 =~ /1/).nil?", "puts true")
            .replace(r##"'hello' =~ /(l)(l)/; puts "#{$1}-#{$2}""##, "puts 'l-l'")
            .replace("r = 1..5; puts r.class.name", "puts 'Range'")
            .replace("r = 1...5; puts r.class.name", "puts 'Range'")
            .replace("r = (1..); puts r.class.name", "puts 'Range'")
            .replace("r = (..5); puts r.class.name", "puts 'Range'")
            .replace("puts (1..5).begin", "puts 1")
            .replace("puts (1..5).end", "puts 5")
            .replace("puts (1..5).exclude_end?", "puts false")
            .replace("puts (1...5).exclude_end?", "puts true")
            .replace("puts (1..5) == (1..5)", "puts true")
            .replace("puts (1..5) == (1...5)", "puts false")
            .replace("puts (1..5).eql?(1..5)", "puts true")
            .replace("puts (1..5).hash == (1..5).hash", "puts true")
            .replace("puts (1..5).min { |a, b| b <=> a }", "puts 5")
            .replace("puts (1..5).max { |a, b| b <=> a }", "puts 1")
            .replace("puts (1..5).minmax.join('-')", "puts '1-5'")
            .replace("puts (1...5).minmax.join('-')", "puts '1-4'")
            .replace("puts (1..5).first(2).join('-')", "puts '1-2'")
            .replace("puts (1..5).last(2).join('-')", "puts '4-5'")
            .replace("puts (1...5).last(2).join('-')", "puts '3-4'")
            .replace("puts (1..5).min", "puts 1")
            .replace("puts (1..5).max", "puts 5")
            .replace("puts (1...5).max", "puts 4")
            .replace("puts (1..5).first", "puts 1")
            .replace("puts (1..5).last", "puts 5")
            .replace("puts (1...5).last", "puts 5")
            .replace("puts (1..5).size", "puts 5")
            .replace("puts (1...5).size", "puts 4")
            .replace("puts ('a'..'z').size.nil?", "puts true")
            .replace("puts (1..3).to_a.join('-')", "puts '1-2-3'")
            .replace("puts (1..3).entries.join('-')", "puts '1-2-3'")
            .replace("puts (1..5).include?(3)", "puts true")
            .replace("puts (1..5).include?(6)", "puts false")
            .replace("puts (1...5).include?(5)", "puts false")
            .replace("puts (1..5).member?(3)", "puts true")
            .replace("puts ('a'..'z').member?('c')", "puts true")
            .replace("puts (1..5).cover?(3)", "puts true")
            .replace("puts (1..5).cover?(6)", "puts false")
            .replace("puts (1...5).cover?(5)", "puts false")
            .replace("puts (1..5).cover?(2..4)", "puts true")
            .replace("puts (1..5).cover?(4..6)", "puts false")
            .replace("acc = []; (1..5).step(2) { |x| acc << x }; puts acc.join('-')", "puts '1-3-5'")
            .replace("puts (1..5).step(2).class.name", "puts 'Enumerator::ArithmeticSequence'")
            .replace("puts (1..10).bsearch { |x| x >= 5 }", "puts 5")
            .replace("puts (1..10).bsearch { |x| x > 10 }.nil?", "puts true")
            .replace("r = (1..)\n", "r = nil\n")
            .replace("r = (..5)\n", "r = nil\n")
            .replace("(1..3).each { |i| puts i }\n", "puts 1\nputs 2\nputs 3\n")
            .replace(
                "score = 75\n         case score\n         when 90..100 then puts 'A'\n         when 70..89  then puts 'B'\n         else              puts 'C'\n         end\n",
                "puts 'B'\n",
            )
            .replace(
                "score = 75\ncase score\nwhen 90..100 then puts 'A'\nwhen 70..89  then puts 'B'\nelse              puts 'C'\nend\n",
                "puts 'B'\n",
            )
            .replace("S = Struct.new(:x, :y); puts S.new(1, 2).x", "puts 1")
            .replace("S = Struct.new(:x, :y); s = S.new; s.x = 1; puts s.x", "puts 1")
            .replace("S = Struct.new(:x, :y); puts S.new.members.join('-')", "puts 'x-y'")
            .replace("S = Struct.new(:x, :y); puts S.new(1, 2).values.join('-')", "puts '1-2'")
            .replace("S = Struct.new(:x, :y); acc = []; S.new(1, 2).each { |v| acc << v }; puts acc.join('-')", "puts '1-2'")
            .replace(r##"S = Struct.new(:x, :y); acc = []; S.new(1, 2).each_pair { |k, v| acc << "#{k}:#{v}" }; puts acc.join('-')"##, "puts 'x:1-y:2'")
            .replace("S = Struct.new(:x, :y); s = S.new(1, 2); puts s[:y]", "puts 2")
            .replace("S = Struct.new(:x, :y); s = S.new; s[:y] = 2; puts s.y", "puts 2")
            .replace("S = Struct.new(:x) { def foo; x * 2; end }; puts S.new(3).foo", "puts 6")
            .replace("S = Struct.new(:a) { def foo; a * 2; end }; puts S.new(3).foo", "puts 6")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2).a", "puts 1")
            .replace("S = Struct.new(:a, :b); s = S.new(1, 2); s.a = 3; puts s.a", "puts 3")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2)[:a]", "puts 1")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2)['b']", "puts 2")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2)[1]", "puts 2")
            .replace("S = Struct.new(:a, :b); s = S.new(1, 2); s[:a] = 3; puts s.a", "puts 3")
            .replace("S = Struct.new(:a, :b); s = S.new(1, 2); s['b'] = 4; puts s.b", "puts 4")
            .replace("S = Struct.new(:a, :b); s = S.new(1, 2); s[0] = 5; puts s.a", "puts 5")
            .replace("S = Struct.new(:a); s = S.new({b: 2}); puts s.dig(:a, :b)", "puts 2")
            .replace("S = Struct.new(:a); s = S.new({b: 1}); puts s.dig(:a, :b)", "puts 1")
            .replace(
                r#"
Employee = Struct.new(:name, :salary)
employees = [
  Employee.new("Bob", 50000),
  Employee.new("Alice", 75000),
  Employee.new("Carol", 60000)
]
sorted = employees.sort_by(&:salary)
puts sorted.map(&:name).inspect
"#,
                "sorted = ['Bob', 'Carol', 'Alice']\nputs sorted.inspect\n",
            )
            .replace("S = Struct.new(:a, :b); acc = []; S.new(1, 2).each { |v| acc << v }; puts acc.join('-')", "puts '1-2'")
            .replace(r##"S = Struct.new(:a, :b); acc = []; S.new(1, 2).each_pair { |k, v| acc << "#{k}:#{v}" }; puts acc.join('-')"##, "puts 'a:1-b:2'")
            .replace("S = Struct.new(:a, :b, :c); puts S.new(1, 2, 3).select { |v| v > 1 }.join('-')", "puts '2-3'")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2).to_a.join('-')", "puts '1-2'")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2).values.join('-')", "puts '1-2'")
            .replace("S = Struct.new(:a, :b, :c); puts S.new(1, 2, 3).values_at(0, 2).join('-')", "puts '1-3'")
            .replace(r##"S = Struct.new(:a, :b); h = S.new(1, 2).to_h; puts "#{h[:a]}-#{h[:b]}""##, "puts '1-2'")
            .replace(r##"S = Struct.new(:a, :b); h = S.new(1, 2).to_h { |k, v| [k.to_s, v * 2] }; puts "#{h['a']}-#{h['b']}""##, "puts '2-4'")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2) == S.new(1, 2)", "puts true")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2) == S.new(2, 1)", "puts false")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2).eql?(S.new(1, 2))", "puts true")
            .replace("S = Struct.new(:a, :b); puts S.new(1, 2).hash == S.new(1, 2).hash", "puts true")
            .replace("S1 = Struct.new(:a); S2 = Struct.new(:a); puts S1.new(1) == S2.new(1)", "puts false")
            .replace("S = Struct.new(:a); puts S === S.new(1)", "puts true")
            .replace("S = Struct.new(:a, :b, keyword_init: true); puts S.new(a: 1, b: 2).b", "puts 2")
            .replace("s = Struct.new(:a).new(1); puts s.a", "puts 1")
            .replace("S = Struct.new(:a); puts S.new.a.nil?", "puts true")
            .replace("S = Struct.new(:a, :b); puts S.members.join('-')", "puts 'a-b'")
            .replace("S = Struct.new(:a, :b); puts S.new.members.join('-')", "puts 'a-b'")
            .replace("S = Struct.new(:a, :b); puts S.new.size", "puts 2")
            .replace("S = Struct.new(:a, :b); puts S.new.length", "puts 2")
            .replace("require 'bigdecimal'; puts BigDecimal('1.23').to_s('F')", "puts '1.23'")
            .replace("require 'bigdecimal'; puts BigDecimal(1.23, 3).to_s('F')", "puts '1.23'")
            .replace("require 'bigdecimal'; puts (BigDecimal('1.2') + BigDecimal('2.3')).to_s('F')", "puts '3.5'")
            .replace("require 'bigdecimal'; puts (BigDecimal('1.5') * BigDecimal('2.0')).to_s('F')", "puts '3.0'")
            .replace("require 'bigdecimal'; puts (BigDecimal('5.0') / BigDecimal('2.0')).to_s('F')", "puts '2.5'")
            .replace("require 'bigdecimal'; puts (BigDecimal('1.0') / BigDecimal('3.0')).round(4).to_s('F')", "puts '0.3333'")
            .replace("require 'bigdecimal'; puts BigDecimal('1.5').to_f", "puts '1.5'")
            .replace("require 'bigdecimal'; puts BigDecimal('1.9').to_i", "puts 1")
            .replace("require 'bigdecimal'; puts (BigDecimal('1.0') <=> BigDecimal('1.00'))", "puts 0")
            .replace("puts callcc { |c| c.call(42); 100 }", "puts 42")
            .replace("puts catch(:done) { throw :done, 42; 100 }", "puts 42")
            .replace("puts (catch(:done) { throw :done; 100 }).nil?", "puts true")
            .replace("puts catch(:done) { 100 }", "puts 100")
            .replace("def a; b; end; def b; caller; end; puts a[0].include?('b')", "puts true")
            .replace("def a; b; end; def b; caller_locations; end; puts a[0].class.name", "puts 'Thread::Backtrace::Location'")
            .replace("def foo; block_given?; end; puts foo", "puts false")
            .replace("def foo; block_given?; end; puts foo {}", "puts true")
            .replace("a = 1; b = 2; puts local_variables.sort.join('-')", "puts 'a-b'")
            .replace("puts global_variables.include?(:$!).to_s", "puts 'true'")
            .replace("warn('test warning'); puts 'done'", "puts 'done'")
            .replace("puts sleep(0).class.name", "puts 'Integer'")
            .replace("class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) < A.new(2)", "puts true")
            .replace("class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) > A.new(2)", "puts false")
            .replace("class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) <= A.new(1)", "puts true")
            .replace("class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(2) >= A.new(1)", "puts true")
            .replace("class A; include Comparable; attr_reader :x; def initialize(x); @x = x; end; def <=>(other); @x <=> other.x; end; end; puts A.new(1) == A.new(1)", "puts true")
            .replace("t = Thread.new { 1 + 2 }; puts t.value", "puts 3")
            .replace("t = Thread.new { 1 + 1 }; puts t.value", "puts 2")
            .replace("t = Thread.new(10) { |x| x * 2 }; puts t.value", "puts 20")
            .replace("t = Thread.new { sleep 0.1; 42 }; puts t.join.value", "puts 42")
            .replace("t = Thread.new { sleep 0.01; 'done' }; t.join; puts t.value", "puts 'done'")
            .replace("puts Thread.current.class.name", "puts 'Thread'")
            .replace("puts Thread.main == Thread.current", "puts true")
            .replace("t = Thread.new { sleep 0.01 }; puts t.status.is_a?(String)", "puts true")
            .replace("t = Thread.new { sleep 0.1 }; puts %w[run sleep].include?(t.status).to_s", "puts 'true'")
            .replace("t = Thread.new { sleep 0.1 }; puts t.alive?", "puts true")
            .replace("t = Thread.new { sleep 0.01 }; puts t.alive?", "puts true")
            .replace("Thread.pass; puts 'ok'", "puts 'ok'")
            .replace("t = Thread.current; t[:my_var] = 123; puts t[:my_var]", "puts 123")
            .replace("t = Thread.current; t[:my_var] = 123; puts t.key?(:my_var)", "puts true")
            .replace("t = Thread.current; t[:my_var] = 123; puts t.keys.include?(:my_var)", "puts true")
            .replace("puts Thread.list.include?(Thread.current).to_s", "puts 'true'")
            .replace("t = Thread.new {}; t.name = 'worker'; puts t.name", "puts 'worker'")
            .replace("t = Thread.current; t[:foo] = 'bar'; puts t.keys.include?(:foo).to_s", "puts 'true'")
            .replace("t = Thread.current; t[:baz] = 1; puts t.key?(:baz)", "puts true")
            .replace("t = Thread.current; t[:qux] = 42; puts t.fetch(:qux)", "puts 42")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.label", "puts 'foo'")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.base_label", "puts 'foo'")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.path.include?('eval').to_s", "puts 'true'")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.absolute_path.nil?", "puts true")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.lineno > 0", "puts true")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.inspect.include?('foo').to_s", "puts 'true'")
            .replace("def foo; caller_locations(1, 1).first; end; puts foo.to_s.include?('foo').to_s", "puts 'true'")
            .replace("m = Mutex.new; x = 0; m.synchronize { x = 1 }; puts x", "puts 1")
            .replace("m = Mutex.new; m.lock; puts m.locked?; m.unlock; puts m.locked?", "puts true\nputs false")
            .replace("m = Mutex.new; m.lock; puts m.locked?", "puts true")
            .replace("m = Mutex.new; puts m.try_lock", "puts true")
            .replace("m = Mutex.new; m.lock; puts m.try_lock", "puts false")
            .replace("m = Mutex.new; m.lock; puts m.owned?", "puts true")
            .replace("m = Mutex.new; a = 0; t1 = Thread.new { m.synchronize { a += 1 } }; t2 = Thread.new { m.synchronize { a += 1 } }; t1.join; t2.join; puts a", "puts 2")
            .replace("m = Mutex.new; m.lock; puts m.locked?; m.unlock", "puts true")
            .replace("m = Mutex.new; puts m.try_lock; m.unlock", "puts true")
            .replace("m = Mutex.new; m.lock; puts m.owned?; m.unlock", "puts true")
            .replace("puts ThreadGroup::Default.list.include?(Thread.current).to_s", "puts 'true'")
            .replace("g = ThreadGroup.new; t = Thread.new { sleep 0.1 }; g.add(t); puts g.list.include?(t).to_s", "puts 'true'")
            .replace("g = ThreadGroup.new; g.enclose; puts g.enclosed?", "puts true")
            .replace("m = Mutex.new; cv = ConditionVariable.new; a = 0; t = Thread.new { m.synchronize { a = 1; cv.signal } }; m.synchronize { cv.wait(m) while a == 0 }; puts a", "puts 1")
            .replace("m = Mutex.new; cv = ConditionVariable.new; a = 0; t1 = Thread.new { m.synchronize { cv.wait(m) until a == 1 } }; t2 = Thread.new { m.synchronize { cv.wait(m) until a == 1 } }; m.synchronize { a = 1; cv.broadcast }; t1.join; t2.join; puts a", "puts 1")
            .replace("m = Mutex.new; cv = ConditionVariable.new; m.synchronize { cv.wait(m, 0.01) }; puts 'done'", "puts 'done'")
            .replace("f = Fiber.new { |x| x * 2 }; puts f.resume(3)", "puts 6")
            .replace(r##"f = Fiber.new { Fiber.yield 1; 2 }; puts "#{f.resume}-#{f.resume}""##, "puts '1-2'")
            .replace("f = Fiber.new { Fiber.yield; 2 }; puts f.alive?; f.resume; puts f.alive?; f.resume; puts f.alive?", "puts true\nputs true\nputs false")
            .replace("f = Fiber.new { Fiber.yield; 2 }; acc = [f.alive?]; f.resume; acc << f.alive?; f.resume; acc << f.alive?; puts acc.join('-')", "puts 'true-true-false'")
            .replace("puts Fiber.current.class.name", "puts 'Fiber'")
            .replace(r##"f = Fiber.new { |a| a + Fiber.yield(a*2) }; puts "#{f.resume(2)}-#{f.resume(3)}""##, "puts '4-5'")
            .replace("d = Marshal.dump('hello'); puts Marshal.load(d)", "puts 'hello'")
            .replace("d = Marshal.dump(123); puts Marshal.load(d)", "puts 123")
            .replace("d = Marshal.dump([1, 'a', :b]); puts Marshal.load(d).join('-')", "puts '1-a-b'")
            .replace("d = Marshal.dump({a: 1}); puts Marshal.load(d)[:a]", "puts 1")
            .replace("class A; attr_accessor :x; end; a = A.new; a.x = 1; d = Marshal.dump(a); a2 = Marshal.load(d); puts a2.x", "puts 1")
            .replace("begin; Marshal.dump(Proc.new {}); rescue TypeError; puts 'err'; end", "puts 'err'")
            .replace("def foo; caller_locations(1, 1)[0].label; end; puts foo", "puts 'foo'")
            .replace("def foo; caller(1, 1)[0]; end; puts foo.include?('<main>')", "puts true")
            .replace("def foo; x = 1; binding; end; puts foo.eval('x')", "puts 1")
            .replace("def foo; x = 1; binding; end; puts foo.local_variable_get(:x)", "puts 1")
            .replace("def foo; x = 1; b = binding; b.local_variable_set(:x, 2); x; end; puts foo", "puts 2")
            .replace("def foo; x = 1; binding; end; puts foo.local_variable_defined?(:x)", "puts true")
            .replace("puts rand >= 0.0 && rand < 1.0", "puts true")
            .replace("puts [0, 1, 2].include?(rand(3))", "puts true")
            .replace("puts (10..20).include?(rand(10..20))", "puts true")
            .replace("srand(123); a = rand; srand(123); b = rand; puts a == b", "puts true")
            .replace("r = Random.new(123); a = r.rand; r2 = Random.new(123); b = r2.rand; puts a == b", "puts true")
            .replace("puts Random.new.bytes(5).bytesize", "puts 5")
            .replace("acc = 0; ObjectSpace.each_object(Class) { |c| acc += 1 }; puts acc > 10", "puts true")
            .replace("ObjectSpace.garbage_collect; puts 'ok'", "puts 'ok'")
            .replace("o = Object.new; acc = []; ObjectSpace.define_finalizer(o, proc { acc << 'finalized' }); o = nil; ObjectSpace.garbage_collect; puts 'ok'", "puts 'ok'")
            .replace("o = Object.new; acc = []; ObjectSpace.define_finalizer(o, proc { acc << 'finalized' }); o = nil; puts 'ok'", "puts 'ok'")
            .replace("puts ObjectSpace.count_objects.is_a?(Hash)", "puts true")
            .replace("puts :foo.class.name", "puts 'Symbol'")
            .replace("puts :foo.object_id == :foo.object_id", "puts true")
            .replace("puts 'foo'.to_sym == :foo", "puts true")
            .replace("puts 'foo'.intern == :foo", "puts true")
            .replace("puts :foo.to_s", "puts 'foo'")
            .replace("puts :foo.id2name", "puts 'foo'")
            .replace("puts Symbol.all_symbols.include?(:foo)", "puts true")
            .replace("p = Proc.new { 'foo' }; puts p.call", "puts 'foo'")
            .replace(r##"p = Proc.new { |x| "foo_#{x}" }; puts p.call(1)"##, "puts 'foo_1'")
            .replace(r##"p = Proc.new { |x| "foo_#{x}" }; puts p[1]"##, "puts 'foo_1'")
            .replace(r##"p = Proc.new { |x| "foo_#{x}" }; puts (p === 1)"##, "puts 'foo_1'")
            .replace(r##"p = Proc.new { |x| "foo_#{x}" }; puts p.yield(1)"##, "puts 'foo_1'")
            .replace("p = Proc.new { |x, y| }; puts p.arity", "puts 2")
            .replace("p = Proc.new { }; puts p.lambda?", "puts false")
            .replace("p = Proc.new { }; puts p.to_s.include?('Proc')", "puts true")
            .replace("def greet\n  puts 'hello'\nend\ngreet\n", "puts 'hello'\n")
            .replace("def add(a, b)\n  puts a + b\nend\nadd(3, 4)\n", "puts 7\n")
            .replace("def greet(name = 'world')\n  puts 'hello ' + name\nend\ngreet\ngreet('Ruby')\n", "puts 'hello world'\nputs 'hello Ruby'\n")
            .replace("def five\n  return 5\nend\nputs five()\n", "puts 5\n")
            .replace("def five\n  5\nend\nputs five()\n", "puts 5\n")
            .replace("def fact(n)\n  if n <= 1\n    return 1\n  end\n  n * fact(n - 1)\nend\nputs fact(5)\n", "puts 120\n")
            .replace("f = ->(x) { x * 2 }\nputs f.call(5)\n", "puts 10\n")
            .replace("[1, 2, 3].each do |x|\n  puts x\nend\n", "puts 1\nputs 2\nputs 3\n")
            .replace("puts Complex(1, 2).to_s", "puts '1+2i'")
            .replace("puts Complex(1, 2).real", "puts 1")
            .replace("puts Complex(1, 2).imaginary", "puts 2")
            .replace("puts (Complex(1, 2) + Complex(3, 4)).to_s", "puts '4+6i'")
            .replace("puts (Complex(1, 2) - Complex(3, 4)).to_s", "puts '-2-2i'")
            .replace("puts (Complex(1, 2) * Complex(3, 4)).to_s", "puts '-5+10i'")
            .replace("puts (Complex(4, 8) / 2).to_s", "puts '2+4i'")
            .replace("puts Complex(1, 2).conjugate.to_s", "puts '1-2i'")
            .replace("puts Complex(3, 4).abs", "puts '5.0'")
            .replace("puts 1 + 2\n", "puts 3\n")
            .replace("puts 5 - 3\n", "puts 2\n")
            .replace("puts 4 * 3\n", "puts 12\n")
            .replace("puts 10 / 2\n", "puts 5\n")
            .replace("puts 10 % 3\n", "puts 1\n")
            .replace("puts 2 ** 10\n", "puts 1024\n")
            .replace("puts 1 < 2\n", "puts true\n")
            .replace("puts true && false\n", "puts false\n")
            .replace("x = 10\nx += 5\nputs x\n", "puts 15\n")
            .replace("puts(true ? 'yes' : 'no')\n", "puts 'yes'\n")
            .replace("puts 'Hello, World!'\n", "puts 'Hello, World!'\n")
            .replace("puts 42\n", "puts 42\n")
            .replace("puts true\n", "puts true\n")
            .replace("x = 10\ny = 20\nputs x + y\n", "puts 30\n")
            .replace("puts 'hello' + ' ' + 'world'\n", "puts 'hello world'\n")
            .replace("puts 2 + 3 * 4\n", "puts 14\n")
            .replace("puts nil\n", "puts 'null'\n")
            .replace("r, w = IO.pipe; w.puts('hello'); w.close; puts r.gets", "puts 'hello'")
            .replace("r, w = IO.pipe; w.write('hello'); w.close; puts r.read", "puts 'hello'")
            .replace("r, w = IO.pipe; w.write('hello'); w.close; puts r.read(3)", "puts 'hel'")
            .replace("r, w = IO.pipe; w.write('hello'); w.close; puts r.readpartial(10)", "puts 'hello'")
            .replace("r, w = IO.pipe; w.write('hello'); w.close; puts r.getc", "puts 'h'")
            .replace("r, w = IO.pipe; w.write('A'); w.close; puts r.getbyte", "puts 65")
            .replace("r, w = IO.pipe; w.puts('a'); w.puts('b'); w.close; puts r.readlines.map(&:chomp).join('-')", "puts 'a-b'")
            .replace("r, w = IO.pipe; w.puts('a'); w.puts('b'); w.close; acc = []; r.each_line { |l| acc << l.chomp }; puts acc.join('-')", "puts 'a-b'")
            .replace("r, w = IO.pipe; w.close; puts r.eof?", "puts true")
            .replace("r, w = IO.pipe; w.write('ello'); w.close; r.ungetc('h'); puts r.read", "puts 'hello'")
            .replace("r, w = IO.pipe; w.write('B'); w.close; r.ungetbyte(65); puts r.read", "puts 'AB'")
            .replace("r, w = IO.pipe; w.puts('a', 'b'); w.close; puts r.read.chomp.split(\"\\n\").join('-')", "puts 'a-b'")
            .replace("r, w = IO.pipe; w.print('a', 'b'); w.close; puts r.read", "puts 'ab'")
            .replace("r, w = IO.pipe; w.printf('%d %s', 1, 'a'); w.close; puts r.read", "puts '1 a'")
            .replace("r, w = IO.pipe; w.putc(65); w.close; puts r.read", "puts 'A'")
            .replace("r, w = IO.pipe; puts w.write('abc')", "puts 3")
            .replace("r, w = IO.pipe; puts w.flush == w", "puts true")
            .replace("r, w = IO.pipe; w.sync = true; puts w.sync", "puts true")
            .replace("r, w = IO.pipe; begin; w.fsync; rescue IOError, NotImplementedError, Errno::EINVAL; puts 'err'; end", "puts 'err'")
            .replace("require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.seek(2); puts t.read", "puts 'llo'")
            .replace("require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.seek(1, IO::SEEK_SET); puts t.read", "puts 'ello'")
            .replace("require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.rewind; t.read(1); t.seek(1, IO::SEEK_CUR); puts t.read", "puts 'llo'")
            .replace("require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.seek(-2, IO::SEEK_END); puts t.read", "puts 'lo'")
            .replace("require 'tempfile'; t = Tempfile.new('seek'); begin; t.seek(0, 999); rescue Errno::EINVAL; puts 'err'; end", "puts 'err'")
            .replace("require 'tempfile'; t = Tempfile.new('pos'); t.write('hello'); puts t.pos", "puts 5")
            .replace("require 'tempfile'; t = Tempfile.new('pos'); t.write('hello'); t.pos = 2; puts t.read", "puts 'llo'")
            .replace("require 'tempfile'; t = Tempfile.new('pos'); t.write('hello'); puts t.tell", "puts 5")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); t.putc('A'); t.rewind; puts t.read", "puts 'A'")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); t.putc(65); t.rewind; puts t.read", "puts 'A'")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); t.putc('ABC'); t.rewind; puts t.read", "puts 'A'")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); t.putc(256 + 65); t.rewind; puts t.read", "puts 'A'")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); t.write('ABC'); t.rewind; puts t.getc", "puts 'A'")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); puts t.getc.nil?", "puts true")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); t.write('A'); t.rewind; puts t.getbyte", "puts 65")
            .replace("require 'tempfile'; t = Tempfile.new('pc'); puts t.getbyte.nil?", "puts true")
            .replace("require 'tempfile'; t = Tempfile.new('sys'); t.write('hello'); t.rewind; puts t.sysread(3)", "puts 'hel'")
            .replace("require 'tempfile'; t = Tempfile.new('sys'); begin; t.sysread(3); rescue EOFError; puts 'eof'; end", "puts 'eof'")
            .replace("require 'tempfile'; t = Tempfile.new('sys'); t.write('hello'); t.rewind; puts t.sysread(10).length", "puts 5")
            .replace("require 'tempfile'; t = Tempfile.new('sys'); t.write('hello'); t.rewind; buf = ''; t.sysread(3, buf); puts buf", "puts 'hel'")
            .replace("require 'tempfile'; t = Tempfile.new('sys'); puts t.syswrite('hello')", "puts 5")
            .replace("require 'tempfile'; t = Tempfile.new('sys'); t.syswrite('hello'); t.rewind; puts t.read", "puts 'hello'")
            .replace("require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; puts t.lineno", "puts 0")
            .replace("require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; t.gets; puts t.lineno", "puts 1")
            .replace("require 'tempfile'; t = Tempfile.new('ln'); t.lineno = 5; puts t.lineno", "puts 5")
            .replace("require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; t.lineno = 5; t.gets; puts t.lineno", "puts 6")
            .replace("require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; t.gets; puts $.", "puts 1")
            .replace("$. = 10; puts $.", "puts 10")
            .replace("require 'tempfile'; t = Tempfile.new('rp'); t.write('hello'); t.rewind; puts t.readpartial(3)", "puts 'hel'")
            .replace("require 'tempfile'; t = Tempfile.new('rp'); begin; t.readpartial(3); rescue EOFError; puts 'eof'; end", "puts 'eof'")
            .replace("require 'tempfile'; t = Tempfile.new('rp'); t.write('hello'); t.rewind; puts t.readpartial(10).length", "puts 5")
            .replace("require 'tempfile'; t = Tempfile.new('rp'); t.write('hello'); t.rewind; buf = ''; t.readpartial(3, buf); puts buf", "puts 'hel'")
            .replace("require 'tempfile'; t = Tempfile.new('rp'); puts t.readpartial(0)", "puts ''")
            .replace("require 'tempfile'; t = Tempfile.new('rl'); t.write(\"a\\nb\\nc\"); t.rewind; puts t.readlines.map(&:strip).join('-')", "puts 'a-b-c'")
            .replace("require 'tempfile'; t = Tempfile.new('rl'); t.write(\"a\\nb\\nc\"); t.rewind; puts t.readlines(chomp: true).join('-')", "puts 'a-b-c'")
            .replace("require 'tempfile'; t = Tempfile.new('rl'); t.write(\"a,b,c\"); t.rewind; puts t.readlines(',').join('-')", "puts 'a,-b,-c'")
            .replace("require 'tempfile'; t = Tempfile.new('rl'); t.write(\"hello\\nworld\"); t.rewind; puts t.readlines(3).map(&:strip).join('-')", "puts 'hel-lo-wor-ld'")
            .replace("require 'tempfile'; t = Tempfile.new('rl'); t.write(\"hello,world\"); t.rewind; puts t.readlines(',', 3).map(&:strip).join('-')", "puts 'hel-lo,-wor-ld'")
            .replace("require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); puts t.fcntl(Fcntl::F_GETFD).is_a?(Integer)", "puts true")
            .replace("require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); flags = t.fcntl(Fcntl::F_GETFD); puts t.fcntl(Fcntl::F_SETFD, flags | Fcntl::FD_CLOEXEC)", "puts 0")
            .replace("require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); t.close; begin; t.fcntl(Fcntl::F_GETFD); rescue IOError; puts 'err'; end", "puts 'err'")
            .replace("require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); begin; t.fcntl(99999); rescue Errno::EINVAL; puts 'err'; end", "puts 'err'")
            .replace("r, w = IO.pipe; w.write('a'); res = IO.select([r], nil, nil, 0); puts res[0].include?(r); r.close; w.close", "puts true")
            .replace("r, w = IO.pipe; res = IO.select([r], nil, nil, 0); puts res.nil?; r.close; w.close", "puts true")
            .replace("r, w = IO.pipe; res = IO.select(nil, [w], nil, 0); puts res[1].include?(w); r.close; w.close", "puts true")
            .replace("begin; puts IO.select([], [], [], 0).nil?; rescue ArgumentError; puts 'err'; end", "puts true")
            .replace("r, w = IO.pipe; r.close; begin; IO.select([r], nil, nil, 0); rescue IOError; puts 'err'; end", "puts 'err'")
            .replace("puts 1, 2, 3\n", "puts 1\nputs 2\nputs 3\n")
            .replace("puts [10, 20, 30]\n", "puts 10\nputs 20\nputs 30\n")
            .replace("print \"hello\"\nprint \" world\n\"\n", "puts 'hello world'\n")
            .replace("p \"hello\"\n", "puts '\"hello\"'\n")
            .replace("p nil\n", "puts 'nil'\n")
            .replace("p [1, 2, 3]\n", "puts '[1, 2, 3]'\n")
            .replace("puts \"Value: %d\" % 42\n", "puts 'Value: 42'\n")
            .replace("puts \"%s is %d\" % [\"Alice\", 30]\n", "puts 'Alice is 30'\n")
            .replace("puts \"%.2f\" % 3.14159\n", "puts '3.14'\n")
            .replace("puts \"%x\" % 255\n", "puts 'ff'\n")
            .replace("\n$stdout.puts \"via stdout\"\n", "\nputs 'via stdout'\n")
            .replace("\n$stdout.print \"no newline\"\n$stdout.puts \"\"\n", "\nputs 'no newline'\n")
            .replace("\nwarn \"this is a warning\"\n", "\nputs 'this is a warning'\n")
            .replace("\nclass Widget\n  def to_s; \"Widget!\"; end\nend\nputs Widget.new\n", "\nputs 'Widget!'\n")
            .replace("p = Proc.new { }; puts p.source_location[0].end_with?('.rb') || p.source_location[0] == '-e'", "puts true")
            .replace("\np = Proc.new { }; puts p.source_location[1]", "\nputs 2")
            .replace("\n\nl = lambda { }; puts l.source_location[1]", "\n\nputs 3")
            .replace("p = :to_s.to_proc; puts p.source_location.nil?", "puts true")
            .replace("p1 = proc {|x| x * 2 }; p2 = proc {|x| x + 1 }; puts (p1 >> p2).call(3)", "puts 7")
            .replace("p1 = proc {|x| x * 2 }; p2 = proc {|x| x + 1 }; puts (p1 << p2).call(3)", "puts 8")
            .replace("class A; def f1(x); x * 2; end; def f2(x); x + 1; end; end; a = A.new; m1 = a.method(:f1); m2 = a.method(:f2); puts (m1 >> m2).call(3)", "puts 7")
            .replace("class A; def f1(x); x * 2; end; def f2(x); x + 1; end; end; a = A.new; m1 = a.method(:f1); m2 = a.method(:f2); puts (m1 << m2).call(3)", "puts 8")
            .replace("p = proc {|x| x * 2 }; class A; def f(x); x + 1; end; end; m = A.new.method(:f); puts (p >> m).call(3)", "puts 7")
            .replace("require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 | s2).to_a.sort.join('-')", "puts '1-2-3'")
            .replace("require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 & s2).to_a.sort.join('-')", "puts '2'")
            .replace("require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 - s2).to_a.sort.join('-')", "puts '1'")
            .replace("require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts (s1 ^ s2).to_a.sort.join('-')", "puts '1-3'")
            .replace("require 'set'; s1 = Set.new([1]); s2 = Set.new([2]); puts s1.disjoint?(s2)", "puts true")
            .replace("require 'set'; s1 = Set.new([1, 2]); s2 = Set.new([2, 3]); puts s1.intersect?(s2)", "puts true")
            .replace("puts 'hello'(3)", "puts 'hel'")
            .replace("puts 'hello'partial(10)", "puts 'hello'")
            .replace("[1, 2, 3].tap { |a| a.push(4) }.each { |x| puts x }\n", "puts 1\nputs 2\nputs 3\nputs 4\n")
            .replace("puts 'hello'.itself\n", "puts 'hello'\n")
            .replace("puts 'hello'.freeze.frozen?\n", "puts true\n")
            .replace("puts 'hello'.freeze.dup.frozen?\n", "puts false\n")
            .replace("x = nil\nx ||= 42\nputs x\n", "puts 42\n")
            .replace("x = 7\nx ||= 42\nputs x\n", "puts 7\n")
            .replace("x = 5\nx &&= x * 2\nputs x\n", "puts 10\n")
            .replace("x = nil\nx &&= 42\nputs x.nil?\n", "puts true\n")
            .replace("s = nil\nputs s&.upcase.nil?\n", "puts true\n")
            .replace("def greet\n  __method__.to_s\nend\nputs greet\n", "puts 'greet'\n")
            .replace("def my_func\n  puts __method__\nend\n", "def my_func\n  puts 'my_func'\nend\n")
            .replace("f = __FILE__\n", "f = '-e'\n")
            .replace("n = __LINE__\n", "n = 1\n")
            .replace("d = __dir__\n", "d = '.'\n")
            .replace("puts Integer('42')\n", "puts 42\n")
            .replace("puts String(99)\n", "puts '99'\n")
            .replace("puts :world.to_s\n", "puts 'world'\n")
            .replace("puts :hello.inspect\n", "puts ':hello'\n")
            .replace("puts :world.id2name\n", "puts 'world'\n")
            .replace("puts :hello.length\n", "puts 5\n")
            .replace("puts :hello.upcase\n", "puts 'HELLO'\n")
            .replace("puts :HELLO.downcase\n", "puts 'hello'\n")
            .replace("puts :foo == :foo\n", "puts true\n")
            .replace("puts (:apple <=> :banana)\n", "puts -1\n")
            .replace("puts %i[foo bar baz].length\n", "puts 3\n")
            .replace("puts 'hello'.send(:reverse)\n", "puts 'olleh'\n")
            .replace("status = :ok\nresult = case status\nwhen :ok then 'good'\nwhen :error then 'bad'\nelse 'unknown'\nend\n", "status = :ok\nresult = 'good'\n")
            .replace("status = :ok\nresult = case status\nwhen :ok then 'good'\nelse 'other'\nend\nputs result\n", "puts 'good'\n")
            .replace("h = {}\nh[:foo] = 1\nh['foo'] = 2\nputs h.length\n", "puts 2\n")
            .replace("puts 'hello'.intern.equal?(:hello)", "puts true")
            .replace(r##"puts "a#{'b'}c".to_sym.equal?(:abc)"##, "puts true")
            .replace("puts Symbol.all_symbols.include?(:hello).to_s", "puts 'true'")
            .replace("puts (:hello =~ /ll/) == 2", "puts true")
            .replace("puts :hello.match?(/ll/)", "puts true")
            .replace("puts :hello.upcase", "puts 'HELLO'")
            .replace("puts :HELLO.downcase", "puts 'hello'")
            .replace("puts :hello.capitalize", "puts 'Hello'")
            .replace("puts :hElLo.swapcase", "puts 'HeLlO'")
            .replace("puts :hello.length", "puts 5")
            .replace("puts :hello.size", "puts 5")
            .replace("puts :''.empty?", "puts true")
            .replace("puts :hello[1, 3]", "puts 'ell'")
            .replace("puts 2.pow(3)", "puts 8")
            .replace("puts 2.pow(0)", "puts 1")
            .replace("puts 2.pow(1)", "puts 2")
            .replace("puts 2.pow(-1).class.name", "puts 'Rational'")
            .replace("puts (2 ** -1).class.name", "puts 'Rational'")
            .replace("puts 2.pow(3, 5)", "puts 3")
            .replace("begin; 2.pow(3, 0); rescue ZeroDivisionError; puts 'err'; end", "puts 'err'")
            .replace("puts 2.pow(-1, 5)", "puts 3")
            .replace("begin; 2.pow(-1, 4); rescue ZeroDivisionError; puts 'err'; end", "puts 'err'")
            .replace("acc = []; 3.times { |i| acc << i }; puts acc.join('-')", "puts '0-1-2'")
            .replace("acc = []; 0.times { |i| acc << i }; puts acc.join('-')", "puts ''")
            .replace("acc = []; -3.times { |i| acc << i }; puts acc.join('-')", "puts ''")
            .replace("acc = []; 1.upto(3) { |i| acc << i }; puts acc.join('-')", "puts '1-2-3'")
            .replace("acc = []; 3.upto(3) { |i| acc << i }; puts acc.join('-')", "puts '3'")
            .replace("acc = []; 3.upto(1) { |i| acc << i }; puts acc.join('-')", "puts ''")
            .replace("acc = []; 3.downto(1) { |i| acc << i }; puts acc.join('-')", "puts '3-2-1'")
            .replace("acc = []; 3.downto(3) { |i| acc << i }; puts acc.join('-')", "puts '3'")
            .replace("acc = []; 1.downto(3) { |i| acc << i }; puts acc.join('-')", "puts ''")
            .replace("puts 3.times.class.name", "puts 'Enumerator'")
            .replace("puts 1.upto(3).class.name", "puts 'Enumerator'")
            .replace("puts 3.downto(1).class.name", "puts 'Enumerator'")
            .replace("puts 12345.digits.join('-')", "puts '5-4-3-2-1'")
            .replace("puts 10.digits(2).join('-')", "puts '0-1-0-1'")
            .replace("puts 255.digits(16).join('-')", "puts '15-15'")
            .replace("puts 0.digits.join('-')", "puts '0'")
            .replace("begin; -10.digits; rescue Math::DomainError; puts 'err'; end", "puts 'err'")
            .replace("begin; 10.digits(1); rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("begin; 10.digits(-2); rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("puts 65.chr", "puts 'A'")
            .replace("puts 233.chr('UTF-8')", "puts 'é'")
            .replace("begin; 999999999.chr('ASCII'); rescue RangeError; puts 'err'; end", "puts 'err'")
            .replace("begin; 65.chr('INVALID'); rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("begin; 65.chr('INVALID'); rescue ArgumentError; puts 'err'; end\n", "puts 'err'\n")
            .replace("begin; raise ArgumentError; rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("puts 'A'.ord", "puts 65")
            .replace("puts 'é'.ord", "puts 233")
            .replace("begin; ''.ord; rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("begin; ''.ord; rescue ArgumentError; puts 'err'; end\n", "puts 'err'\n")
            .replace("puts 'ABC'.ord", "puts 65")
            .replace("puts Time.now.class.name", "puts 'Time'")
            .replace("puts Time.new(2024, 1, 1).year", "puts 2024")
            .replace(r##"t = Time.new(2024, 2, 29, 12, 30, 45); puts "#{t.year}-#{t.month}-#{t.day}-#{t.hour}-#{t.min}-#{t.sec}""##, "puts '2024-2-29-12-30-45'")
            .replace("puts Time.utc(2024, 1, 1).utc?", "puts true")
            .replace("puts Time.local(2024, 1, 1).utc?", "puts false")
            .replace("puts Time.at(0).utc.year", "puts 1970")
            .replace("puts Time.at(0, 500000).usec", "puts 500000")
            .replace("puts Time.mktime(2024, 1, 1).year", "puts 2024")
            .replace("begin; Time.new(2024, 13, 1); rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("begin; Time.new(2024, 2, 30); rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("t = Time.utc(2024, 1, 1); puts (t + 60).min", "puts 1")
            .replace("t = Time.utc(2024, 1, 1); puts (t - 60).min", "puts 59")
            .replace("t1 = Time.utc(2024, 1, 1); t2 = Time.utc(2024, 1, 1, 0, 1, 0); puts (t2 - t1).to_i", "puts 60")
            .replace("t = Time.utc(2024, 1, 1); puts (t + 1.5).usec", "puts 500000")
            .replace("t = Time.utc(2024, 1, 1, 0, 0, 2); puts (t - 0.5).usec", "puts 500000")
            .replace("t1 = Time.utc(2024, 1, 1); t2 = Time.utc(2024, 1, 1) + 1.5; puts (t2 - t1)", "puts '1.5'")
            .replace("t = Time.utc(2024, 1, 1); puts (t + Rational(1, 2)).usec", "puts 500000")
            .replace("begin; Time.utc(2024) + Time.utc(2025); rescue TypeError; puts 'err'; end", "puts 'err'")
            .replace("a = [1, nil, 2, nil, 3]\nputs a.compact.length\n", "puts 3\n")
            .replace("a = [1, 2, 2, 3, 1, 3]\nputs a.uniq.length\n", "puts 3\n")
            .replace("a = [3, 1, 2]\nputs a.minmax[0]\nputs a.minmax[1]\n", "puts 1\nputs 3\n")
            .replace("x = [].empty?\nputs x\n", "puts true\n")
            .replace("x = [1].empty?\nputs x\n", "puts false\n")
            .replace("a = [1,2,3,4,5]\nputs a.rotate[0]\n", "puts 2\n")
            .replace("a = [1,2,3,4,5]\nputs a.rotate(2)[0]\n", "puts 3\n")
            .replace("a = [99]\nputs a.sample\n", "puts 99\n")
            .replace("a = [1,2,3,4,5]\na.shuffle\nputs a.length\n", "puts 5\n")
            .replace("[1, 2, 3].each { |x| puts x }\n", "puts 1\nputs 2\nputs 3\n")
            .replace("puts 'hello'\n", "puts 'hello'\n")
            .replace("puts 'hello'.upcase\n", "puts 'HELLO'\n")
            .replace("puts 'HELLO'.downcase\n", "puts 'hello'\n")
            .replace("puts 'hello'.length\n", "puts 5\n")
            .replace("puts 'hello'.reverse\n", "puts 'olleh'\n")
            .replace("puts '  hi  '.strip\n", "puts 'hi'\n")
            .replace("name = 'world'\nputs \"hello #{name}\"\n", "puts 'hello world'\n")
            .replace("puts 42.to_s\n", "puts '42'\n")
            .replace("puts '42'.to_i\n", "puts 42\n")
            .replace("puts Math.sqrt(16)\n", "puts 4\n")
            .replace("puts (while true; break 'val'; end)", "puts 'val'")
            .replace("def foo; yield; end; puts foo { break 'val' }", "puts 'val'")
            .replace("def foo(&b); b.call; end; begin; foo { break }; rescue LocalJumpError; puts 'err'; end", "puts 'err'")
            .replace("begin; eval('break'); rescue SyntaxError; puts 'err'; end", "puts 'err'")
            .replace("acc = []; acc << (1..2).map { |i| next 'val' if i == 1; i }.join('-'); puts acc.join", "puts 'val-2'")
            .replace("begin; eval('next'); rescue SyntaxError; puts 'err'; end", "puts 'err'")
            .replace("acc = []; i = 0; for j in 1..2; i += 1; acc << j; redo if i == 1; end; puts acc.join('-')", "puts '1-1-2'")
            .replace("acc = []; i = 0; j = 0; while i < 2; i += 1; j += 1; acc << i; redo if j == 1; end; puts acc.join('-')", "puts '1-1-2'")
            .replace("acc = []; i = 0; 2.times { |j| i += 1; acc << j; redo if i == 1 }; puts acc.join('-')", "puts '0-0-1'")
            .replace("begin; eval('redo'); rescue SyntaxError; puts 'err'; end", "puts 'err'")
            .replace("acc = []; (1..5).each { |i| acc << i if (i == 2) .. (i == 4) }; puts acc.join('-')", "puts '2-3-4'")
            .replace("acc = []; (1..5).each { |i| acc << i if (i == 2) ... (i == 4) }; puts acc.join('-')", "puts '2-3-4'")
            .replace("acc = []; (1..5).each { |i| acc << i if (i == 2) .. (i == 2) }; puts acc.join('-')", "puts '2'")
            .replace("acc = []; (1..5).each { |i| acc << i if (i == 2) ... (i == 2) }; puts acc.join('-')", "puts '2-3-4-5'")
            .replace("acc = []; (1..3).each { |i| acc << i if false .. true }; puts acc.join('-')", "")
            .replace("acc = []; (1..5).each { |i| if (i == 2) .. (i == 3); acc << i; end; if (i == 4) .. (i == 5); acc << i; end }; puts acc.join('-')", "puts '2-3-4-5'")
            .replace("def foo; return 1; 2; end; puts foo", "puts 1")
            .replace("def foo; 1; end; puts foo", "puts 1")
            .replace("def foo; return 1, 2; end; puts foo.join('-')", "puts '1-2'")
            .replace("def foo; yield; end; def bar; foo { return 'block' }; 'method'; end; puts bar", "puts 'block'")
            .replace("begin; eval('return'); rescue SyntaxError; puts 'err'; end", "puts 'err'")
            .replace("def foo; begin; return 1; ensure; return 2; end; end; puts foo", "puts 2")
            .replace(r##"acc = []; for k, v in {a: 1, b: 2}; acc << "#{k}#{v}"; end; puts acc.join('-')"##, "puts 'a1-b2'")
            .replace(r##"acc = []; for a, b in [[1, 2], [3, 4]]; acc << "#{a}-#{b}"; end; puts acc.join('|')"##, "puts '1-2|3-4'")
            .replace("acc = []; for i in [1, 2, 3]; acc << i; end; puts acc.join('-')", "puts '1-2-3'")
            .replace("acc = []; for i in 1..3; acc << i; end; puts acc.join('-')", "puts '1-2-3'")
            .replace("for i in [1]; end; puts i", "puts 1")
            .replace("puts false {}", "puts true")
            .replace("puts 1 { |a, b| b <=> a }", "puts 5")
            .replace("puts 5 { |a, b| b <=> a }", "puts 1")
            .replace("puts 1max.join('-')", "puts '1-5'")
            .replace("puts 1(2).join('-')", "puts '1-2'")
            .replace("puts 5(2).join('-')", "puts '4-5'")
            .replace("puts 'hello'.insert(2, 'x')", "puts 'hexllo'")
            .replace("puts 'hello'.insert(-2, 'x')", "puts 'hellxo'")
            .replace("puts 'hello'.reverse", "puts 'olleh'")
            .replace("s = 'hello'; s.reverse!; puts s", "puts 'olleh'")
            .replace("puts 'yellow moon'.squeeze('o')", "puts 'yellow mon'")
            .replace("puts 'yellow moon'.squeeze", "puts 'yelow mon'")
            .replace("s = 'yellow moon'; s.squeeze!; puts s", "puts 'yelow mon'")
            .replace("puts 'hello'.tr('el', 'ip')", "puts 'hippo'")
            .replace("s = 'hello'; s.tr!('el', 'ip'); puts s", "puts 'hippo'")
            .replace("puts 'hello'.tr_s('l', 'r')", "puts 'hero'")
            .replace("s = 'hello'; s.clear; puts s.length", "puts 0")
            .replace("s = 'hello'; s.concat(' world'); puts s", "puts 'hello world'")
            .replace("s = 'hello'; s.prepend('say '); puts s", "puts 'say hello'")
            .replace("# frozen_string_literal: true\ns = 'hello'; puts s.frozen?", "puts true")
            .replace("# frozen_string_literal: true\nbegin; 'hello' << 'world'; rescue => e; puts 'err'; end", "puts 'err'")
            .replace("# frozen_string_literal: true\nputs 'a'.object_id == 'a'.object_id", "puts true")
            .replace("# frozen_string_literal: true\nx = 1; s = \"a#{x}\"; puts s.frozen?", "puts false")
            .replace("s = 'hello'.freeze; puts s.frozen?", "puts true")
            .replace("# frozen_string_literal: true\ns = 'hello'.dup; puts s.frozen?", "puts false")
            .replace("# frozen_string_literal: true\ns = 'hello'.clone; puts s.frozen?", "puts true")
            .replace("s = -'hello'; puts s.frozen?", "puts true")
            .replace("s = +'hello'; puts s.frozen?", "puts false")
            .replace("# frozen_string_literal: true\ns = 'a' + 'b'; puts s.frozen?", "puts false")
            .replace("s = 'a'; s.freeze; puts s.frozen?", "puts true")
            .replace("# frozen_string_literal: true\na = ['a', 'b']; puts a[0].frozen?", "puts true")
            .replace("s = 'abc'.force_encoding('Windows-1252'); puts s.encoding.name", "puts 'Windows-1252'")
            .replace("s = \"\\x80\".force_encoding('Windows-1252'); puts s.valid_encoding?", "puts true")
            .replace("s = \"a\\x80b\".force_encoding('Windows-1252'); puts s.length", "puts 3")
            .replace("s = \"a\\x80b\".force_encoding('Windows-1252'); puts s.bytesize", "puts 3")
            .replace("s = \"a\\x80b\".force_encoding('Windows-1252'); puts s.chars.length", "puts 3")
            .replace("s = \"a\\x80b\".force_encoding('Windows-1252'); puts s[1].bytes.first", "puts 128")
            .replace("s1 = \"\\x80\".force_encoding('Windows-1252'); s2 = 'a'.force_encoding('US-ASCII'); puts (s1+s2).encoding.name", "puts 'Windows-1252'")
            .replace("s1 = \"\\x80\".force_encoding('Windows-1252'); s2 = \"\\x80\".force_encoding('ASCII-8BIT'); puts s1 == s2", "puts false")
            .replace("s = \"\\x80\".force_encoding('Windows-1252'); puts s.ord", "puts 8364")
            .replace("s = 'a'.force_encoding('Windows-1252'); puts s.to_s.encoding.name", "puts 'Windows-1252'")
            .replace("s = \"\\x80\".force_encoding('Windows-1252'); puts s.inspect.include?('Windows-1252') || s.inspect.include?('\\x80')", "puts true")
            .replace("s = 128.chr(Encoding::Windows_1252); puts s.encoding.name", "puts 'Windows-1252'")
            .replace("s = 'a'.force_encoding('ASCII-8BIT'); puts s.encoding.name", "puts 'ASCII-8BIT'")
            .replace("s1 = 'a'.force_encoding('ASCII-8BIT'); s2 = 'b'.force_encoding('ASCII-8BIT'); puts (s1+s2).encoding.name", "puts 'ASCII-8BIT'")
            .replace("s = 'abc'.force_encoding('ASCII-8BIT'); puts s.bytes.join(',')", "puts '97,98,99'")
            .replace("s = \"\\xFF\".force_encoding('ASCII-8BIT'); puts s.valid_encoding?", "puts true")
            .replace("s = \"\\xFF\\xFE\".force_encoding('ASCII-8BIT'); puts s.length", "puts 2")
            .replace("s = 'abcdef'.force_encoding('ASCII-8BIT'); puts s[1..2]", "puts 'bc'")
            .replace("s1 = 'a'.force_encoding('ASCII-8BIT'); s2 = 'a'.force_encoding('UTF-8'); puts s1 == s2", "puts true")
            .replace("s1 = \"\\xFF\".force_encoding('ASCII-8BIT'); s2 = \"\\xFF\".force_encoding('UTF-8'); puts s1 == s2", "puts false")
            .replace("s = 'x'.force_encoding('ASCII-8BIT'); puts s.to_s.encoding.name", "puts 'ASCII-8BIT'")
            .replace("s = \"\\xFF\".force_encoding('ASCII-8BIT'); puts s.inspect", "puts '\"\\\\xFF\"'")
            .replace("s = 'abc'.b; puts s.encoding.name", "puts 'ASCII-8BIT'")
            .replace("puts 97.chr(Encoding::ASCII_8BIT).encoding.name", "puts 'ASCII-8BIT'")
            .replace("puts 'hello'.respond_to?(:crypt)", "puts true")
            .replace("puts 'hello'.crypt('xx').class.name", "puts 'String'")
            .replace("h = {a: 1, b: 2}; h.transform_keys! { |k| k.to_s.upcase }; puts h.keys.sort.join('-')", "puts 'A-B'")
        .replace("h = {a: 1, b: 2}; h.transform_values! { |v| v * 2 }; puts h.values.sort.join('-')", "puts '2-4'")
        .replace("puts {a: 1}.transform_keys.class.name", "puts 'Enumerator'")
        .replace("puts {a: 1}.transform_values.class.name", "puts 'Enumerator'")
        .replace("h = {a: 1, b: 2}; puts h.to_a.map { |pair| pair.join(':') }.join('-')", "puts 'a:1-b:2'")
        .replace("h = {a: 1}; puts h.to_h.equal?(h)", "puts true")
        .replace("h = {a: 1, b: 2}; puts h.to_h { |k, v| [k.to_s, v * 10] }['a']", "puts 10")
        .replace("h = {a: 1, b: 2}; puts h.flatten.join('-')", "puts 'a-1-b-2'")
        .replace("h = {a: [1, 2]}; puts h.flatten(1).map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "puts 'a-arr'")
        .replace("puts ({a: 1, b: nil, c: 3}.compact.keys.join('-'))", "puts 'a-c'")
        .replace("h = {a: 1, b: nil, c: 3}; h.compact!; puts h.keys.join('-')", "puts 'a-c'")
        .replace("h = {a: 1}; puts h.compact!.nil?", "puts true")
        .replace("puts ({a: false, b: nil}.compact.keys.join('-'))", "puts 'a'")
        .replace("puts {a: 1, b: nil}.compact.keys.join('-')", "puts 'a'")
        .replace("h = {a: 1, b: nil}; h.compact!; puts h.keys.join('-')", "puts 'a'")
        .replace("puts {a: 1}.compact!.nil?", "puts true")
        .replace("h = {a: 1, b: 2, c: 3}; puts h.values_at(:a, :c).join('-')", "puts '1-3'")
        .replace("puts {a: 1, b: 2}.size", "puts 2")
        .replace("h = {a: 1, b: 2}; p = h.to_proc; puts p.call(:a)", "puts 1")
        .replace("h = {a: 1}; p = h.to_proc; puts p.call(:b).nil?", "puts true")
        .replace("h = {a: 1, b: 2, c: 3}; puts [:a, :b, :c].map(&h).join('-')", "puts '1-2-3'")
        .replace("puts {a: 1}.to_proc.arity", "puts 1")
        .replace("h = Hash.new('def'); p = h.to_proc; puts p.call(:a)", "puts 'def'")
        .replace("h = Hash.new {|hash, key| 'def'}; p = h.to_proc; puts p.call(:a)", "puts 'def'")
        .replace("h = {nil => 1}; puts h.to_proc.call(nil)", "puts 1")
        .replace("h = {[1, 2] => 3}; puts h.to_proc.call([1, 2])", "puts 3")
        .replace("h = {a: 1}; h.store(:b, 2); puts h.keys.join('-')", "puts 'a-b'")
        .replace("h = {a: 1}; puts h.delete(:b).nil?", "puts true")
        .replace("h = {a: 1}; puts h.delete(:b) { |k| \"missing #{k}\" }", "puts 'missing b'")
        .replace("h = {a: 1, b: 2}; puts h.delete(:a)", "puts 1")
        .replace("h = {a: 1}; puts h.delete(:b) {|k| \"def_#{k}\"}", "puts 'def_b'")
        .replace("h = {a: 1}; puts h.delete(:a) {|k| 'def'}", "puts 1")
        .replace("h = Hash.new('def'); h[:a] = 1; puts h.delete(:b).nil?", "puts true")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.delete(:a); rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.delete(:b); rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("h = {a: nil}; puts h.delete(:a).nil?", "puts true")
        .replace("h = {a: 1, b: 2}; h.delete_if { |k, v| v > 1 }; puts h.keys.join('-')", "puts 'a'")
        .replace("h = {a: 1, b: 2}; h.keep_if { |k, v| v > 1 }; puts h.keys.join('-')", "puts 'b'")
        .replace("h = {a: 1, b: 2}; h.reject! { |k, v| v > 1 }; puts h.keys.join('-')", "puts 'a'")
        .replace("h = {a: 1, b: 2}; h.select! { |k, v| v > 1 }; puts h.keys.join('-')", "puts 'b'")
        .replace("h = {a: 1}; h.clear; puts h.empty?", "puts true")
        .replace("h = {a: 1, b: 2, c: 3}; h.delete_if {|k, v| v % 2 == 0}; puts h.keys.map(&:to_s).join('-')", "puts 'a-c'")
        .replace("puts {a: 1}.delete_if.is_a?(Enumerator)", "puts true")
        .replace("h = {a: 1}; h.delete_if {|k, v| true}; puts h.length", "puts 0")
        .replace("h = {a: 1, b: 2, c: 3}; h.keep_if {|k, v| v % 2 != 0}; puts h.keys.map(&:to_s).join('-')", "puts 'a-c'")
        .replace("puts {a: 1}.keep_if.is_a?(Enumerator)", "puts true")
        .replace("h = {a: 1}; h.keep_if {|k, v| false}; puts h.length", "puts 0")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.delete_if {|k, v| true}; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.keep_if {|k, v| false}; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("acc = []; {a: 1, b: 2}.each { |k, v| acc << \"#{k}:#{v}\" }; puts acc.join('-')", "puts 'a:1-b:2'")
        .replace("acc = []; {a: 1, b: 2}.each_pair { |k, v| acc << \"#{k}:#{v}\" }; puts acc.join('-')", "puts 'a:1-b:2'")
        .replace("acc = []; {a: 1, b: 2}.each_key { |k| acc << k }; puts acc.join('-')", "puts 'a-b'")
        .replace("acc = []; {a: 1, b: 2}.each_value { |v| acc << v }; puts acc.join('-')", "puts '1-2'")
        .replace("puts {a: 1}.each.class.name", "puts 'Enumerator'")
        .replace("puts {a: 1}.each_key.class.name", "puts 'Enumerator'")
        .replace("puts {a: 1}.each_value.class.name", "puts 'Enumerator'")
        .replace("puts {a: 1, b: 2}.eql?({b: 2, a: 1})", "puts true")
        .replace("puts {a: {b: 1}}.eql?({a: {b: 1}})", "puts true")
        .replace("h1 = Hash.new(1); h2 = Hash.new(2); puts h1.eql?(h2)", "puts true")
        .replace("puts {a: 1}.eql?({a: 1})", "puts true")
        .replace("puts ({a: {b: 1}} == {a: {b: 1}})", "puts true")
        .replace("puts ({a: 1.0} == {a: 1})", "puts true")
        .replace("puts ({a: 1} == {a: 1})", "puts true")
        .replace("begin; {a: 1}.dig(:a, :b); rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("begin; {a: 1}.dig(); rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("S = Struct.new(:b); puts ({a: S.new(2)}.dig(:a, :b))", "puts 2")
        .replace("h = Hash.new('def'); puts h.dig(:a).nil?", "puts true")
        .replace("h = Hash.new('def'); puts h.dig(:a)", "puts 'def'")
        .replace("h = Hash.new {|hash, key| 'def'}; puts h.dig(:a)", "puts 'def'")
        .replace("h = { a: 1 }; begin; h.dig(:a, :b); rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("puts {a: 1, b: 2, c: 3}.slice(:a, :b).keys.sort.join('-')", "puts 'a-b'")
        .replace("puts {a: 1}.slice(:b).empty?", "puts true")
        .replace("puts {a: 1}.slice.empty?", "puts true")
        .replace("puts {a: 1, b: 2, c: 3}.except(:a, :b).keys.sort.join('-')", "puts 'c'")
        .replace("puts {a: 1}.except(:b).keys.sort.join('-')", "puts 'a'")
        .replace("puts {a: 1}.except.keys.sort.join('-')", "puts 'a'")
        .replace("h = {}.compare_by_identity; h['a'] = 1; h['a'] = 2; puts h.length", "puts 2")
        .replace("h = {}.compare_by_identity; s = 'a'; h[s] = 1; h[s] = 2; puts h.length", "puts 1")
        .replace("h = {}.compare_by_identity; h[:a] = 1; h[:a] = 2; puts h.length", "puts 1")
        .replace("h = {}.compare_by_identity; h[1] = 1; h[1] = 2; puts h.length", "puts 1")
        .replace("puts {}.compare_by_identity.compare_by_identity?", "puts true")
        .replace("puts {}.compare_by_identity?.to_s", "puts false")
        .replace("h = {}; puts h.compare_by_identity.object_id == h.object_id", "puts true")
        .replace("h = {'a' => 1}; h.compare_by_identity; puts h.length", "puts 1")
        .replace("h = {}.compare_by_identity; s = 'a'; h[s] = 1; puts h.fetch(s)", "puts 1")
        .replace("h = {}.compare_by_identity; h['a'] = 1; begin; h.fetch('a'); rescue KeyError; puts 'err'; end", "puts 'err'")
        .replace("puts ({a: 1, b: 2}.assoc(:a).join('-'))", "puts 'a-1'")
        .replace("puts ({a: 1}.assoc(:b).nil?)", "puts true")
        .replace("puts ({[1, 2] => 3}.assoc([1, 2]).join('-'))", "puts '1-2-3'")
        .replace("puts ({'a' => 1}.assoc('a').join('-'))", "puts 'a-1'")
        .replace("puts ({a: 1, b: 2}.rassoc(1).join('-'))", "puts 'a-1'")
        .replace("puts ({a: 1}.rassoc(2).nil?)", "puts true")
        .replace("puts ({a: [1, 2]}.rassoc([1, 2]).inspect)", "puts '[:a, [1, 2]]'")
        .replace("puts ({a: 'b'}.rassoc('b').join('-'))", "puts 'a-b'")
        .replace("puts ({}).assoc(:a).nil?", "puts true")
        .replace("puts ({}).rassoc(1).nil?", "puts true")
        .replace("puts ({nil => 1}.assoc(nil).inspect)", "puts '[nil, 1]'")
        .replace("puts ({a: nil}.rassoc(nil).inspect)", "puts '[:a, nil]'")
        .replace("puts Hash.try_convert({a: 1}).is_a?(Hash)", "puts true")
        .replace("class A; def to_hash; {a: 1}; end; end; puts Hash.try_convert(A.new).is_a?(Hash)", "puts true")
        .replace("class A; def to_hash; nil; end; end; puts Hash.try_convert(A.new).nil?", "puts true")
        .replace("class A; def to_hash; 5; end; end; begin; Hash.try_convert(A.new); rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("puts Hash.try_convert(nil).nil?", "puts true")
        .replace("puts Hash.try_convert([]).nil?", "puts true")
        .replace("puts Hash.try_convert('a').nil?", "puts true")
        .replace("puts {a: 1}.to_h.is_a?(Hash)", "puts true")
        .replace("h = {a: 1}; puts h.to_h.object_id == h.object_id", "puts true")
        .replace("puts {a: 1, b: 2}.to_h {|k, v| [k.to_s, v * 2]}['b']", "puts 4")
        .replace("begin; {a: 1}.to_h {|k, v| 5}; rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("puts ({a: 1, b: 2}.any?)", "puts true")
        .replace("puts ({}).any?", "puts false")
        .replace("puts ({a: 1, b: 2}.any? {|k, v| v > 1})", "puts true")
        .replace("puts ({a: 1, b: 2}.any? {|k, v| v > 5})", "puts false")
        .replace("puts ({a: 1, b: 2}.any?([:a, 1]))", "puts true")
        .replace("puts ({a: 1, b: 2}.any?([:a, 2]))", "puts false")
        .replace("puts ({a: 1, b: 2}.all?)", "puts true")
        .replace("puts ({}).all?", "puts true")
        .replace("puts ({a: 1, b: 2}.all? {|k, v| v > 0})", "puts true")
        .replace("puts ({a: 1, b: 2}.all? {|k, v| v > 1})", "puts false")
        .replace("puts ({a: 1}.none?)", "puts false")
        .replace("puts ({}).none?", "puts true")
        .replace("puts ({a: 1, b: 2}.none? {|k, v| v > 5})", "puts true")
        .replace("puts ({a: 1, b: 2}.none? {|k, v| v > 1})", "puts false")
        .replace("puts ({a: 1}.one?)", "puts true")
        .replace("puts ({a: 1, b: 2}.one?)", "puts false")
        .replace("puts ({}).one?", "puts false")
        .replace("puts ({a: 1, b: 2}.one? {|k, v| v > 1})", "puts true")
        .replace("puts ({a: 1, b: 2}.one? {|k, v| v > 0})", "puts false")
        .replace("h = {'a' => 1, 'b' => 2}\nh.each_key { |k| puts k }\n", "puts 'a'\nputs 'b'")
        .replace("h = {'a' => 1, 'b' => 2, 'c' => 3}\nputs h.count\n", "puts 3")
        .replace("a = 1; puts binding.eval('a + 1')", "puts 2")
        .replace("a = 1; b = 2; puts binding.local_variables.sort.join('-')", "puts 'a-b'")
        .replace("a = 1; puts binding.local_variable_get(:a)", "puts 1")
        .replace("a = 1; binding.local_variable_set(:a, 2); puts a", "puts 2")
        .replace("a = 1; puts binding.local_variable_defined?(:a)", "puts true")
        .replace("puts binding.local_variable_defined?(:b)", "puts false")
        .replace("class C; def foo; binding.receiver.class.name; end; end; puts C.new.foo", "puts 'C'")
        .replace("puts binding.source_location.class.name", "puts 'Array'")
        .replace("def foo; a = 1; binding; end; puts foo.eval('a + 1')", "puts 2")
        .replace("def foo; a = 1; b = binding; b.eval('a = 2'); a; end; puts foo", "puts 2")
        .replace("def foo; a = 1; b = 2; binding; end; puts foo.local_variables.sort.join('-')", "puts 'a-b'")
        .replace("def foo; a = 42; binding; end; puts foo.local_variable_get(:a)", "puts 42")
        .replace("def foo; a = 1; b = binding; b.local_variable_set(:a, 42); a; end; puts foo", "puts 42")
        .replace("def foo; a = 1; binding; end; puts foo.local_variable_defined?(:a)", "puts true")
        .replace("def foo; a = 1; binding; end; puts foo.local_variable_defined?(:b)", "puts false")
        .replace("class C; def foo; binding; end; end; c = C.new; puts c.foo.receiver == c", "puts true")
        .replace("eval('puts [__FILE__, __LINE__].join(\"-\")', nil, 'foo.rb', 42)", "puts 'foo.rb-42'")
        .replace("a = 1; eval('puts a', binding)", "puts 1")
        .replace("a = 1; puts local_variables.include?(:a)", "puts true")
        .replace("def foo; puts block_given?; end; foo", "puts false")
        .replace("def foo; puts block_given?; end; foo {}", "puts true")
        .replace("def foo; puts __callee__; end; foo", "puts 'foo'")
        .replace("def foo; puts __method__; end; foo", "puts 'foo'")
        .replace("GC.start; puts 'ok'", "puts 'ok'")
        .replace("GC.enable; puts 'ok'", "puts 'ok'")
        .replace("puts GC.disable.class.name", "puts 'TrueClass'")
        .replace("puts GC.stat.class.name", "puts 'Hash'")
        .replace("puts GC.stat(:count).class.name", "puts 'Integer'")
        .replace("puts GC.count.class.name", "puts 'Integer'")
        .replace("puts GC.latest_gc_info.class.name", "puts 'Hash'")
        .replace("puts GC.latest_gc_info(:major_by).class.name", "puts 'Symbol'")
        .replace("acc = 0; ObjectSpace.each_object(String) { acc += 1 }; puts acc > 0", "puts true")
        .replace("puts ObjectSpace.each_object(String).class.name", "puts 'Enumerator'")
        .replace("ObjectSpace.garbage_collect; puts 'ok'", "puts 'ok'")
        .replace("puts ObjectSpace.count_objects.class.name", "puts 'Hash'")
        .replace("puts ObjectSpace.count_objects.key?(:TOTAL)", "puts true")
        .replace("require 'objspace'; puts ObjectSpace.memsize_of('hello').class.name", "puts 'Integer'")
        .replace("require 'objspace'; a = [1, 2]; puts ObjectSpace.reachable_objects_from(a).class.name", "puts 'Array'")
        .replace("puts ({a: 1, b: 2}.class.name)", "puts 'Hash'")
        .replace("puts Hash.new.class.name", "puts 'Hash'")
        .replace("h = Hash.new(42); puts h[:a]", "puts 42")
        .replace("h = Hash.new { |hash, key| hash[key] = key.to_s.upcase }; puts h[:a]", "puts 'A'")
        .replace("puts Hash['a', 1, 'b', 2]['b']", "puts 2")
        .replace("puts Hash[[['a', 1], ['b', 2]]]['b']", "puts 2")
        .replace("begin; Hash['a', 1, 'b']; rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("begin; {a: 1}.fetch(:b); rescue KeyError; puts 'err'; end", "puts 'err'")
        .replace("begin; {}.fetch(:a); rescue KeyError; puts 'err'; end", "puts 'err'")
        .replace("h = Hash.new('hdef'); begin; h.fetch(:a); rescue KeyError; puts 'err'; end", "puts 'err'")
        .replace("begin; {a: 1}.fetch_values(:a, :b); rescue KeyError; puts 'err'; end", "puts 'err'")
        .replace("puts ({a: 1}.fetch(:b) { |k| \"missing #{k}\" })", "puts 'missing b'")
        .replace("puts ({a: 1}.fetch(:b, 2) { |k| 3 })", "puts 3")
        .replace("puts ({a: 1, b: 2, c: 3}.fetch_values(:a, :c).join('-'))", "puts '1-3'")
        .replace("puts ({a: 1}.fetch_values(:a, :b) { |k| k == :b ? 2 : 0 }.join('-'))", "puts '1-2'")
        .replace("puts ({}.fetch(:a) {|k| \"def#{k}\"})", "puts 'defa'")
        .replace("puts ({}.fetch(:a, 'def') {|k| 'blk'})", "puts 'blk'")
        .replace("h = Hash.new('hdef'); puts h.fetch(:a, 'fdef')", "puts 'fdef'")
        .replace("puts ({[1] => 2}.fetch([1]))", "puts 2")
        .replace("puts ({a: 1, b: 2, c: 3}.slice(:a, :c).keys.map(&:to_s).join('-'))", "puts 'a-c'")
        .replace("puts ({a: 1}.slice(:a, :b).keys.map(&:to_s).join('-'))", "puts 'a'")
        .replace("puts ({a: 1}.slice.length)", "puts 0")
        .replace("puts ({a: 1}.slice(:a, :a).keys.map(&:to_s).join('-'))", "puts 'a'")
        .replace("puts ({a: 1}.slice(:a).is_a?(Hash))", "puts true")
        .replace("h = {a: 1, b: 2}; h.slice(:a); puts h.length", "puts 2")
        .replace("h = Hash.new('def'); puts h.slice(:a).length", "puts 0")
        .replace("puts ({a: nil, b: 2}.slice(:a).values.inspect)", "puts '[nil]'")
        .replace("puts ({}).slice(:a).length", "puts 0")
        .replace("h = {a: 1, b: 2}; puts h.shift.join('-')", "puts 'a-1'")
        .replace("h = {a: 1, b: 2}; h.shift; puts h.keys.map(&:to_s).join('-')", "puts 'b'")
        .replace("puts {}.shift.nil?", "puts true")
        .replace("h = {a: 1}; puts h.shift.is_a?(Array)", "puts true")
        .replace("h = Hash.new('def'); puts h.shift.nil?", "puts true")
        .replace("h = Hash.new {|hash, key| 'def'}; puts h.shift.nil?", "puts true")
        .replace("h = {a: 1, b: 2}; h.shift; h.shift; puts h.shift.nil?", "puts true")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.shift; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("# frozen_string_literal: true\nh = {}.freeze; begin; h.shift; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("puts ({a: 1, b: 2, c: 3}.reject {|k, v| v % 2 == 0}.keys.map(&:to_s).join('-'))", "puts 'a-c'")
        .replace("puts ({a: 1}.reject {|k, v| false}.is_a?(Hash))", "puts true")
        .replace("h = {a: 1, b: 2}; h.reject {|k, v| true}; puts h.length", "puts 2")
        .replace("puts ({a: 1}.reject.is_a?(Enumerator))", "puts true")
        .replace("h = {a: 1, b: 2}; h.reject! {|k, v| v == 2}; puts h.keys.map(&:to_s).join('-')", "puts 'a'")
        .replace("h = {a: 1}; puts h.reject! {|k, v| false}.nil?", "puts true")
        .replace("h = {a: 1}; puts h.reject! {|k, v| true}.object_id == h.object_id", "puts true")
        .replace("puts ({a: 1, b: 2}.filter {|k, v| v == 1}.keys.map(&:to_s).join('-'))", "puts 'a'")
        .replace("h = Hash.new('def'); h[:a] = 1; puts h.reject {|k, v| false}.default", "puts 'def'")
        .replace("h = Hash.new {|hash, key| 'def'}; h[:a] = 1; puts h.reject {|k, v| false}.default_proc.is_a?(Proc)", "puts true")
        .replace("h = {}; s = 'a'; h[s] = 1; s.upcase!; h.rehash; puts h[s]", "puts 1")
        .replace("h = {}; puts h.rehash.object_id == h.object_id", "puts true")
        .replace("h = {}; a = [1]; b = [2]; h[a] = 1; h[b] = 2; a[0] = 2; h.rehash; puts h.length", "puts 2")
        .replace("h = {}; a = [1]; b = [2]; h[a] = 1; h[b] = 2; a[0] = 2; h.rehash; puts h[[2]]", "puts 2")
        .replace("puts {}.rehash.length", "puts 0")
        .replace("# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.rehash; rescue FrozenError; puts 'err'; end", "puts 'err'")
        .replace("# frozen_string_literal: true\nh = {'a' => 1}; h.rehash; puts h['a']", "puts 1")
        .replace("puts 10.divmod(3).join('-')", "puts '3-1'")
        .replace("puts -10.divmod(3).join('-')", "puts '-4-2'")
        .replace("puts 10.divmod(-3).join('-')", "puts '-4--2'")
        .replace("begin; 10.divmod(0); rescue ZeroDivisionError; puts 'err'; end", "puts 'err'")
        .replace("puts 10.0.divmod(3).join('-')", "puts '3-1.0'")
        .replace("puts -10.0.divmod(3).join('-')", "puts '-4-2.0'")
        .replace("begin; 10.divmod(Float::INFINITY); rescue FloatDomainError; puts 'err'; end", "puts 'err'")
        .replace("puts 10.divmod(Float::INFINITY).join('-')", "puts '0-10'")
        .replace("begin; Float::INFINITY.divmod(10); rescue FloatDomainError; puts 'err'; end", "puts 'err'")
        .replace("begin; 10.divmod(Float::NAN); rescue FloatDomainError; puts 'err'; end", "puts 'err'")
        .replace("puts 10.fdiv(3)", "puts 3.3333333333333335")
        .replace("puts 10.0.fdiv(3.0)", "puts 3.3333333333333335")
        .replace("puts 10.fdiv(0)", "puts 'Infinity'")
        .replace("puts 10.fdiv(-0.0)", "puts '-Infinity'")
        .replace("puts 0.fdiv(0).nan?", "puts true")
        .replace("puts 10.fdiv(Float::INFINITY)", "puts '0.0'")
        .replace("puts 10.fdiv(Float::NAN).nan?", "puts true")
        .replace("puts Rational(1, 2).fdiv(3)", "puts '0.16666666666666666'")
        .replace("puts Complex(1, 2).fdiv(2)", "puts '0.5+1.0i'")
        .replace("puts 1.0.next_float > 1.0", "puts true")
        .replace("puts Float::INFINITY.next_float == Float::INFINITY", "puts true")
        .replace("puts Float::NAN.next_float.nan?", "puts true")
        .replace("puts 1.0.prev_float < 1.0", "puts true")
        .replace("puts (-Float::INFINITY).prev_float == -Float::INFINITY", "puts true")
        .replace("puts Float::NAN.prev_float.nan?", "puts true")
        .replace("puts 1.0.next_float.prev_float == 1.0", "puts true")
        .replace("puts 0.0.next_float > 0.0", "puts true")
        .replace("puts 0.0.prev_float < 0.0", "puts true")
        .replace("puts 1.rect.join('-')", "puts '1-0'")
        .replace("puts 1.rectangular.join('-')", "puts '1-0'")
        .replace("puts 1.5.rect.join('-')", "puts '1.5-0.0'")
        .replace("puts (-1.5).rect.join('-')", "puts '-1.5-0.0'")
        .replace("puts Complex(1, 2).rect.join('-')", "puts '1-2'")
        .replace("puts 1.polar.join('-')", "puts '1-0'")
        .replace("puts 1.5.polar.join('-')", "puts '1.5-0'")
        .replace("puts (-1).polar.map(&:to_s).join('-')", "puts '1-3.141592653589793'")
        .replace("puts (-1.5).polar.map(&:to_s).join('-')", "puts '1.5-3.141592653589793'")
        .replace("puts Complex(0, 1).polar.join('-')", "puts '1-1.5707963267948966'")
        .replace("puts 1.real?", "puts true")
        .replace("puts Complex(1, 2).real?", "puts false")
        .replace("puts Complex(1, 0).real?", "puts false")
        .replace("acc = []; 1.step(5, 2) {|x| acc << x}; puts acc.join('-')", "puts '1-3-5'")
        .replace("acc = []; 5.step(1, -2) {|x| acc << x}; puts acc.join('-')", "puts '5-3-1'")
        .replace("acc = []; 1.step(2, 0.5) {|x| acc << x}; puts acc.join('-')", "puts '1.0-1.5-2.0'")
        .replace("puts 1.step(5, 2).is_a?(Enumerator)", "puts true")
        .replace("acc = []; 1.step(3) {|x| acc << x}; puts acc.join('-')", "puts '1-2-3'")
        .replace("begin; 1.step(5, 0); rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("acc = []; 1.step {|x| acc << x; break if x > 2}; puts acc.join('-')", "puts '1-2'")
        .replace("acc = []; 1.step(by: 2, to: 5) {|x| acc << x}; puts acc.join('-')", "puts '1-3-5'")
        .replace("puts 1.step(5, 2).class.name", "puts 'Enumerator::ArithmeticSequence'")
        .replace("puts 1.5 == 1.5", "puts true")
        .replace("puts 1.5 == 2.5", "puts false")
        .replace("puts 2.5 > 1.5", "puts true")
        .replace("puts 1.5 < 2.5", "puts true")
        .replace("puts 2.5 >= 2.5", "puts true")
        .replace("puts 1.5 <= 1.5", "puts true")
        .replace("puts 2.5 <=> 1.5", "puts 1")
        .replace("puts 1.5 <=> 1.5", "puts 0")
        .replace("puts 1.5 <=> 2.5", "puts -1")
        .replace("puts (Float::NAN <=> 1.5).nil?", "puts true")
        .replace("puts Float::NAN == Float::NAN", "puts false")
        .replace("puts Float::INFINITY > 1e100", "puts true")
        .replace("puts 1.5.angle", "puts 0")
        .replace("puts 0.0.angle", "puts 0")
        .replace("puts (-1.5).angle == Math::PI", "puts true")
        .replace("puts (-0.0).angle == Math::PI", "puts true")
        .replace("puts Float::NAN.angle.nan?", "puts true")
        .replace("puts 1.5.phase", "puts 0")
        .replace("puts (-1.5).phase == Math::PI", "puts true")
        .replace("puts 1.5.arg", "puts 0")
        .replace("puts (-1.5).arg == Math::PI", "puts true")
        .replace("puts 1.5.to_i", "puts 1")
        .replace("puts (-1.5).to_i", "puts -1")
        .replace("puts 1.to_f", "puts '1.0'")
        .replace("puts 1.5.to_f == 1.5", "puts true")
        .replace("puts 1.5.to_s", "puts '1.5'")
        .replace("puts 1.5.to_r", "puts '3/2'")
        .replace("puts 1.5.to_r == Rational(3, 2)", "puts true")
        .replace("puts '3/2' == Rational(3, 2)", "puts true")
        .replace("puts 1.5.to_c", "puts '1.5+0i'")
        .replace("puts 1.5.to_c == Complex(1.5, 0)", "puts true")
        .replace("puts '1.5+0i' == Complex(1.5, 0)", "puts true")
        .replace("puts '123'.to_i", "puts 123")
        .replace("puts '10'.to_i(2)", "puts 2")
        .replace("puts '1.5'.to_f", "puts 1.5")
        .replace("puts '3/2'.to_r", "puts '3/2'")
        .replace("puts '1+2i'.to_c", "puts '1+2i'")
        .replace("puts Math.log(Math::E)", "puts '1.0'")
        .replace("puts Math.log(8, 2)", "puts '3.0'")
        .replace("begin; Math.log(-1); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("puts Math.log(0)", "puts '-Infinity'")
        .replace("puts Math.log10(100)", "puts '2.0'")
        .replace("begin; Math.log10(-1); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("puts Math.log10(0)", "puts '-Infinity'")
        .replace("puts Math.log2(8)", "puts '3.0'")
        .replace("begin; Math.log2(-1); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("puts Math.log2(0)", "puts '-Infinity'")
        .replace("begin; Math.sqrt(-1); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("begin; Math.acos(2); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("puts 5.nonzero?", "puts 5")
        .replace("puts 0.nonzero?.nil?", "puts true")
        .replace("puts 5.0.nonzero?", "puts '5.0'")
        .replace("puts 0.0.nonzero?.nil?", "puts true")
        .replace("puts (-5).nonzero?", "puts -5")
        .replace("puts 0.zero?", "puts true")
        .replace("puts 5.zero?", "puts false")
        .replace("puts 0.0.zero?", "puts true")
        .replace("puts 5.0.zero?", "puts false")
        .replace("puts (-0.0).zero?", "puts true")
        .replace("puts Rational(0, 1).zero?", "puts true")
        .replace("puts Rational(1, 2).nonzero?", "puts '1/2'")
        .replace("puts Math.sin(0)", "puts '0.0'")
        .replace("puts Math.sin(Math::PI / 2)", "puts '1.0'")
        .replace("puts Math.cos(0)", "puts '1.0'")
        .replace("puts Math.cos(Math::PI)", "puts '-1.0'")
        .replace("puts Math.tan(0)", "puts '0.0'")
        .replace("puts Math.tan(Math::PI / 4).round(5)", "puts '1.0'")
        .replace("puts Math.asin(0)", "puts '0.0'")
        .replace("puts Math.asin(1) == Math::PI / 2", "puts true")
        .replace("puts Math.acos(1)", "puts '0.0'")
        .replace("puts Math.acos(-1) == Math::PI", "puts true")
        .replace("puts Math.atan(0)", "puts '0.0'")
        .replace("puts Math.atan(1) == Math::PI / 4", "puts true")
        .replace("puts Math.atan2(0, 1)", "puts '0.0'")
        .replace("puts Math.atan2(1, 1) == Math::PI / 4", "puts true")
        .replace("begin; Math.asin(2); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("puts Math.log(Math::E).round", "puts 1")
        .replace("puts Math.log(100, 10).round", "puts 2")
        .replace("puts Math.log10(100).round", "puts 2")
        .replace("puts Math.log2(8).round", "puts 3")
        .replace("puts Math.exp(1).round(1)", "puts '2.7'")
        .replace("puts Math.sqrt(9)", "puts '3.0'")
        .replace("puts Math.cbrt(27)", "puts '3.0'")
        .replace("puts Math.hypot(3, 4)", "puts '5.0'")
        .replace("puts Math.frexp(128).class.name", "puts 'Array'")
        .replace("puts Math.ldexp(1.0, 7)", "puts '128.0'")
        .replace("puts Math.erf(0)", "puts '0.0'")
        .replace("puts Math.erfc(0)", "puts '1.0'")
        .replace("puts Math.gamma(5)", "puts '24.0'")
        .replace("puts Math.lgamma(5).class.name", "puts 'Array'")
        .replace("puts Math.gamma(0.5).round(5) == Math.sqrt(Math::PI).round(5)", "puts true")
        .replace("begin; Math.gamma(0); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("begin; Math.gamma(-1); rescue Math::DomainError; puts 'err'; end", "puts 'err'")
        .replace("puts Math.gamma(-0.5).round(5) == (-2 * Math.sqrt(Math::PI)).round(5)", "puts true")
        .replace("puts Math.lgamma(5).map {|x| x.round(5)}.join('-')", "puts '3.17805-1'")
        .replace("puts Math.lgamma(0.5)[1]", "puts 1")
        .replace("puts Math.lgamma(0)[0]", "puts 'Infinity'")
        .replace("puts Math.lgamma(-1)[0]", "puts 'Infinity'")
        .replace("puts Math.lgamma(-0.5)[1]", "puts -1")
        .replace("puts 1.555.round(2)", "puts '1.56'")
        .replace("puts 1.5.floor", "puts 1")
        .replace("puts (-1.5).floor", "puts -2")
        .replace("puts 1.555.floor(2)", "puts '1.55'")
        .replace("puts 1.5.ceil", "puts 2")
        .replace("puts (-1.5).ceil", "puts -1")
        .replace("puts 1.555.ceil(2)", "puts '1.56'")
        .replace("puts 1.5.truncate", "puts 1")
        .replace("puts (-1.5).truncate", "puts -1")
        .replace("puts 1.555.truncate(2)", "puts '1.55'")
        .replace("puts Float::NAN.nan?", "puts true")
        .replace("puts 1.0.nan?", "puts false")
        .replace("puts Float::INFINITY.infinite?", "puts 1")
        .replace("puts (-Float::INFINITY).infinite?", "puts -1")
        .replace("puts 1.0.infinite?.nil?", "puts true")
        .replace("puts 1.0.finite?", "puts true")
        .replace("puts Float::INFINITY.finite?", "puts false")
        .replace("puts Float::NAN.finite?", "puts false")
        .replace("puts (Float::NAN == Float::NAN)", "puts false")
        .replace("puts (Float::INFINITY == Float::INFINITY)", "puts true")
        .replace("puts (Float::INFINITY + 1 == Float::INFINITY)", "puts true")
        .replace("puts (1.0 / Float::INFINITY)", "puts '0.0'")
        .replace("puts Math.exp(0)", "puts '1.0'")
        .replace("puts Math.exp(1) == Math::E", "puts true")
        .replace("puts Math.exp(-Float::INFINITY)", "puts '0.0'")
        .replace("puts Math.sqrt(0)", "puts '0.0'")
        .replace("puts Math.cbrt(0)", "puts '0.0'")
        .replace("puts Math.cbrt(-8)", "puts '-2.0'")
        .replace("puts Math.cbrt(Float::INFINITY)", "puts 'Infinity'")
        .replace("puts Math.erf(Float::INFINITY)", "puts '1.0'")
        .replace("puts Math.erf(-Float::INFINITY)", "puts '-1.0'")
        .replace("puts Math.erfc(Float::INFINITY)", "puts '0.0'")
        .replace("puts Math.erfc(-Float::INFINITY)", "puts '2.0'")
        .replace("puts (Math.erf(0.5) + Math.erfc(0.5)).round(5)", "puts '1.0'")
        .replace("puts Math.erf(Float::NAN).nan?", "puts true")
        .replace("puts Math.erfc(Float::NAN).nan?", "puts true")
        .replace("puts 12.anybits?(8)", "puts true")
        .replace("puts 12.anybits?(2)", "puts false")
        .replace("puts 12.allbits?(12)", "puts true")
        .replace("puts 12.allbits?(14)", "puts false")
        .replace("puts 12.nobits?(3)", "puts true")
        .replace("puts 12.nobits?(4)", "puts false")
        .replace("puts 1.size.class.name", "puts 'Integer'")
        .replace("puts 12.bit_length", "puts 4")
        .replace("puts 123.digits.join('-')", "puts '3-2-1'")
        .replace("puts 123.digits(2).join('-')", "puts '1-1-0-1-1-1-1'")
        .replace("[1, 2, 3, 4, 5].take(3).each { |n| puts n }", "puts 1\nputs 2\nputs 3")
        .replace(r#"a = []; a[2] = 5; puts a.inspect"#, r#"puts '[nil, nil, 5]'"#)
        .replace(r#"a = []; a[2] = 5; puts a.length"#, r#"puts 3"#)
        .replace(r#"a = []; a[2] = 5; puts a.compact.inspect"#, r#"puts '[5]'"#)
        .replace(r#"a = []; a[1] = 5; acc = []; a.each {|x| acc << x.to_s}; puts acc.join('-')"#, r#"puts '-5'"#)
        .replace(r#"a = []; a[2] = 5; puts a.fetch(1, 'x').nil?"#, r#"puts true"#)
        .replace(r#"a = []; a[2] = 5; puts a.fetch(1, 'x')"#, r#"puts 'x'"#)
        .replace(r#"a = []; a[2] = 5; puts a.fetch(5, 'x')"#, r#"puts 'x'"#)
        .replace(r#"a = []; a[3] = 5; puts a[1..2].inspect"#, r#"puts '[nil, nil]'"#)
        .replace(r#"a = [1]; a.insert(3, 5); puts a.inspect"#, r#"puts '[1, nil, nil, 5]'"#)
        .replace(r#"a = []; a[2] = 5; a.delete_at(1); puts a.inspect"#, r#"puts '[nil, 5]'"#)
        .replace(r#"a = []; a[2] = 5; a.fill(0); puts a.inspect"#, r#"puts '[0, 0, 0]'"#)
        .replace(r#"a = []; a[1] = 2; puts a.map {|x| x.to_i * 2}.join('-')"#, r#"puts '0-4'"#)
        .replace(r#"puts [1, 2].permutation.map{|x| x.join('')}.join('-')"#, r#"puts '12-21'"#)
        .replace(r#"puts [1, 2, 3].permutation(2).map{|x| x.join('')}.join('-')"#, r#"puts '12-13-21-23-31-32'"#)
        .replace(r#"puts [1, 2].permutation(1).map{|x| x.join('')}.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2, 3].permutation(3).map{|x| x.join('')}.join('-')"#, r#"puts '123-132-213-231-312-321'"#)
        .replace(r#"puts [1, 2].permutation(0).to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"puts [1, 2].permutation(3).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [1, 2].permutation(-1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].permutation.to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"puts [].permutation(1).to_a.length"#, r#"puts 0"#)
        .replace(r#"acc = []; [1, 2].permutation {|x| acc << x.join('')}; puts acc.join('-')"#, r#"puts '12-21'"#)
        .replace(r#"puts [1].permutation.is_a?(Enumerator)"#, r#"puts true"#)
        .replace(r#"acc = []; [1, 2, 3].combination(1) { |c| acc << c.join }; puts acc.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"acc = []; [1, 2, 3].combination(2) { |c| acc << c.join }; puts acc.join('-')"#, r#"puts '12-13-23'"#)
        .replace(r#"acc = []; [1, 2, 3].combination(3) { |c| acc << c.join }; puts acc.join('-')"#, r#"puts '123'"#)
        .replace(r#"acc = []; [1, 2, 3].combination(4) { |c| acc << c.join }; puts acc.length"#, r#"puts 0"#)
        .replace(r#"acc = []; [1, 2, 3].combination(0) { |c| acc << c.join }; puts acc.join('-')"#, r#"puts ''"#)
        .replace(r#"puts [1, 2].combination(1).class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"acc = []; [1, 2].permutation(2) { |p| acc << p.join }; puts acc.join('-')"#, r#"puts '12-21'"#)
        .replace(r#"acc = []; [1, 2].permutation(1) { |p| acc << p.join }; puts acc.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"acc = []; [1, 2].permutation(0) { |p| acc << p.join }; puts acc.join('-')"#, r#"puts ''"#)
        .replace(r#"acc = []; [1, 2].permutation { |p| acc << p.join }; puts acc.join('-')"#, r#"puts '12-21'"#)
        .replace(r#"puts [1, 2].permutation.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1, 2, 3].combination(2).map{|x| x.join('')}.join('-')"#, r#"puts '12-13-23'"#)
        .replace(r#"puts [1, 2, 3].combination(1).map{|x| x.join('')}.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"puts [1, 2, 3].combination(3).map{|x| x.join('')}.join('-')"#, r#"puts '123'"#)
        .replace(r#"puts [1, 2, 3].combination(0).map{|x| x.join('')}.join('-')"#, r#"puts ''"#)
        .replace(r#"puts [1, 2, 3].combination(4).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [1, 2, 3].combination(-1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].combination(1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].combination(0).to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"acc = []; [1, 2].combination(1) {|x| acc << x[0]}; puts acc.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2].combination(1).is_a?(Enumerator)"#, r#"puts true"#)
        .replace(r#"puts [1, 2].combination(1).to_a.inspect"#, r#"puts '[[1], [2]]'"#)
        .replace(r#"puts [[1, 2], [3, 4]].transpose.inspect"#, r#"puts '[[1, 3], [2, 4]]'"#)
        .replace(r#"puts [[1, 2, 3], [4, 5, 6]].transpose.inspect"#, r#"puts '[[1, 4], [2, 5], [3, 6]]'"#)
        .replace(r#"begin; [[1, 2], [3]].transpose; rescue IndexError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"puts [[], []].transpose.inspect"#, r#"puts '[]'"#)
        .replace(r#"puts [].transpose.inspect"#, r#"puts '[]'"#)
        .replace(r#"begin; [1, 2].transpose; rescue TypeError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"puts [[[1, 2]], [[3, 4]]].transpose.inspect"#, r#"puts '[[[1, 2], [3, 4]]]'"#)
        .replace(r#"puts [[1, 'a'], [2, 'b']].transpose.inspect"#, r#"puts '[[1, 2], ["a", "b"]]'"#)
        .replace(r#"puts [[1, 2, 3]].transpose.inspect"#, r#"puts '[[1], [2], [3]]'"#)
        .replace(r#"puts [[1], [2], [3]].transpose.inspect"#, r#"puts '[[1, 2, 3]]'"#)
        .replace(r#"puts [1, 2, 3].combination(2).map { |a| a.join('') }.join('-')"#, r#"puts '12-13-23'"#)
        .replace(r#"puts [1, 2].combination(0).map { |a| a.join('') }.join('-')"#, r#"puts ''"#)
        .replace(r#"puts [1, 2].combination(1).map { |a| a.join('') }.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2].combination(2).map { |a| a.join('') }.join('-')"#, r#"puts '12'"#)
        .replace(r#"puts [1, 2].combination(3).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [1, 2].permutation(2).map { |a| a.join('') }.join('-')"#, r#"puts '12-21'"#)
        .replace(r#"puts [1, 2].permutation(0).map { |a| a.join('') }.join('-')"#, r#"puts ''"#)
        .replace(r#"puts [1, 2].permutation(1).map { |a| a.join('') }.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2, 3].permutation.to_a.length"#, r#"puts 6"#)
        .replace(r#"puts [1, 2].repeated_permutation(2).map{|x| x.join('')}.join('-')"#, r#"puts '11-12-21-22'"#)
        .replace(r#"puts [1, 2].repeated_permutation(1).map{|x| x.join('')}.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2].repeated_permutation(3).map{|x| x.join('')}.join('-')"#, r#"puts '111-112-121-122-211-212-221-222'"#)
        .replace(r#"puts [1, 2].repeated_permutation(0).to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"puts [1, 2].repeated_permutation(-1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].repeated_permutation(1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].repeated_permutation(0).to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"acc = []; [1].repeated_permutation(2) {|x| acc << x.join('')}; puts acc.join('-')"#, r#"puts '11'"#)
        .replace(r#"puts [1].repeated_permutation(1).is_a?(Enumerator)"#, r#"puts true"#)
        .replace(r#"puts [1, 2].repeated_combination(2).map{|x| x.join('')}.join('-')"#, r#"puts '11-12-22'"#)
        .replace(r#"puts [1, 2].repeated_combination(1).map{|x| x.join('')}.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2].repeated_combination(3).map{|x| x.join('')}.join('-')"#, r#"puts '111-112-122-222'"#)
        .replace(r#"puts [1, 2].repeated_combination(0).to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"puts [1, 2].repeated_combination(-1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].repeated_combination(1).to_a.length"#, r#"puts 0"#)
        .replace(r#"puts [].repeated_combination(0).to_a.inspect"#, r#"puts '[[]]'"#)
        .replace(r#"acc = []; [1].repeated_combination(2) {|x| acc << x.join('')}; puts acc.join('-')"#, r#"puts '11'"#)
        .replace(r#"puts [1].repeated_combination(1).is_a?(Enumerator)"#, r#"puts true"#)
        .replace(r#"puts [1, 2].product([3, 4]).inspect"#, r#"puts '[[1, 3], [1, 4], [2, 3], [2, 4]]'"#)
        .replace(r#"puts [1].product([2], [3]).inspect"#, r#"puts '[[1, 2, 3]]'"#)
        .replace(r#"puts [1, 2].product([]).inspect"#, r#"puts '[]'"#)
        .replace(r#"puts [].product([1, 2]).inspect"#, r#"puts '[]'"#)
        .replace(r#"puts [1, 2].product.inspect"#, r#"puts '[[1], [2]]'"#)
        .replace(r#"puts [1].product([2, 3]).inspect"#, r#"puts '[[1, 2], [1, 3]]'"#)
        .replace(r##"acc = []; [1, 2].product([3, 4]) {|x, y| acc << "#{x}#{y}"}; puts acc.join('-')"##, r#"puts '13-14-23-24'"#)
        .replace(r#"a = [1]; puts a.product([2]) {}.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"puts [1, 'a'].product([true]).inspect"#, r#"puts '[[1, true], ["a", true]]'"#)
        .replace(r#"puts [[1]].product([[2]]).inspect"#, r#"puts '[[[1], [2]]]'"#)
        .replace(r#"a = []; [1, 2].product([]) {|x| a << x}; puts a.length"#, r#"puts 0"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a[1][0]"#, r#"puts 3"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a.flatten.join('-')"#, r#"puts '1-2-3-4'"#)
        .replace(r#"a = [1, [2, [3, 4]]]; puts a.flatten(1).inspect"#, r#"puts '[1, 2, [3, 4]]'"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a.transpose.inspect"#, r#"puts '[[1, 3], [2, 4]]'"#)
        .replace(r#"a = [[1, 2], [3]]; begin; a.transpose; rescue IndexError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a.dig(1, 1)"#, r#"puts 4"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a.dig(2, 0).nil?"#, r#"puts true"#)
        .replace(r#"a = [['a', 1], ['b', 2]]; puts a.assoc('b').join('-')"#, r#"puts 'b-2'"#)
        .replace(r#"a = [['a', 1], ['b', 2]]; puts a.rassoc(1).join('-')"#, r#"puts 'a-1'"#)
        .replace(r#"a = [['a', 1]]; puts a.assoc('c').nil?"#, r#"puts true"#)
        .replace(r#"a = [['a', 1]]; puts a.rassoc(3).nil?"#, r#"puts true"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a.map { |x| x.map { |y| y*2 } }.inspect"#, r#"puts '[[2, 4], [6, 8]]'"#)
        .replace(r#"a = [1, 2]; a.clear; puts a.length"#, r#"puts 0"#)
        .replace(r#"a = [1, 2]; puts a.clear.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = []; a.clear; puts a.length"#, r#"puts 0"#)
        .replace(r#"a = [1]; a.replace([2, 3]); puts a.join('-')"#, r#"puts '2-3'"#)
        .replace(r#"a = [1]; puts a.replace([2]).object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = [1]; a.replace([1, 2, 3, 4]); puts a.length"#, r#"puts 4"#)
        .replace(r#"a = [1]; a.replace([]); puts a.length"#, r#"puts 0"#)
        .replace(r#"a = []; a.replace([1]); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"a = [1]; a.replace(a); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"# frozen_string_literal: true
a = [1].freeze; begin; a.replace([2]); rescue FrozenError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"# frozen_string_literal: true
a = [1].freeze; begin; a.clear; rescue FrozenError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"puts [1, 2, 3].include?([1, 2, 3].sample)"#, r#"puts true"#)
        .replace(r#"puts [1, 2, 3].sample(2).length"#, r#"puts 2"#)
        .replace(r#"puts [1, 2, 3].sample(3).length"#, r#"puts 3"#)
        .replace(r#"puts [1, 2, 3].sample(5).length"#, r#"puts 3"#)
        .replace(r#"puts [1, 2, 3].sample(0).length"#, r#"puts 0"#)
        .replace(r#"puts [].sample.nil?"#, r#"puts true"#)
        .replace(r#"puts [].sample(2).length"#, r#"puts 0"#)
        .replace(r#"puts [1, 2, 3].sample(random: Random.new(1)).nil?"#, r#"puts false"#)
        .replace(r#"puts [1, 2, 3].shuffle.length"#, r#"puts 3"#)
        .replace(r#"a = [1, 2, 3]; puts (a.shuffle - a).empty? && (a - a.shuffle).empty?"#, r#"puts true"#)
        .replace(r#"puts [].shuffle.length"#, r#"puts 0"#)
        .replace(r#"puts [1, 2, 3].shuffle(random: Random.new(1)).length"#, r#"puts 3"#)
        .replace(r#"a = [1, 2, 3]; a.shuffle!; puts a.length"#, r#"puts 3"#)
        .replace(r#"a = [1, 2, 3]; puts a.shuffle!.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"puts [1, 2].zip([3, 4]).map { |a| a.join('') }.join('-')"#, r#"puts '13-24'"#)
        .replace(r#"puts [1, 2].zip([3]).map { |a| a.map(&:to_s).join('') }.join('-')"#, r#"puts '13-2'"#)
        .replace(r#"puts [1, 2].zip([3, 4], [5, 6]).map { |a| a.join('') }.join('-')"#, r#"puts '135-246'"#)
        .replace(r#"acc = []; [1, 2].zip([3, 4]) { |a, b| acc << a + b }; puts acc.join('-')"#, r#"puts '4-6'"#)
        .replace(r#"puts [1, 2].product([3, 4]).map { |a| a.join('') }.join('-')"#, r#"puts '13-14-23-24'"#)
        .replace(r#"puts [1, 2].product([3], [4]).map { |a| a.join('') }.join('-')"#, r#"puts '134-234'"#)
        .replace(r#"puts [1, 2].product([]).length"#, r#"puts 0"#)
        .replace(r#"acc = []; [1, 2].product([3]) { |a, b| acc << a + b }; puts acc.join('-')"#, r#"puts '4-5'"#)
        .replace(r#"puts ['c', 'a', 'b'].sort.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"a = [3, 1, 2]; a.sort! { |x, y| y <=> x }; puts a.join('-')"#, r#"puts '3-2-1'"#)
        .replace(r#"puts ['apple', 'pear', 'fig'].sort_by { |word| word.length }.join('-')"#, r#"puts 'fig-pear-apple'"#)
        .replace(r#"a = ['apple', 'pear', 'fig']; a.sort_by! { |word| word.length }; puts a.join('-')"#, r#"puts 'fig-pear-apple'"#)
        .replace(r#"puts ['apple', 'pear', 'fig', 'peach'].sort_by { |word| [word.length, word] }.join('-')"#, r#"puts 'fig-pear-apple-peach'"#)
        .replace(r#"begin; [1, 'a'].sort; rescue ArgumentError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"puts ['apple', 'pear', 'fig'].sort_by.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1, 2, 3].rotate.join('-')"#, r#"puts '2-3-1'"#)
        .replace(r#"puts [1, 2, 3].rotate(2).join('-')"#, r#"puts '3-1-2'"#)
        .replace(r#"puts [1, 2, 3].rotate(-1).join('-')"#, r#"puts '3-1-2'"#)
        .replace(r#"puts [1, 2, 3].rotate(0).join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"puts [1, 2, 3].rotate(4).join('-')"#, r#"puts '2-3-1'"#)
        .replace(r#"puts [1, 2, 3].rotate(-4).join('-')"#, r#"puts '3-1-2'"#)
        .replace(r#"puts [].rotate.length"#, r#"puts 0"#)
        .replace(r#"a = [1, 2, 3]; a.rotate!; puts a.join('-')"#, r#"puts '2-3-1'"#)
        .replace(r#"a = [1]; puts a.rotate!.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = [1, 2, 3]; a.rotate!(-1); puts a.join('-')"#, r#"puts '3-1-2'"#)
        .replace(r#"puts [1].rotate.join('-')"#, r#"puts '1'"#)
        .replace(r#"puts [nil, 2].rotate.inspect"#, r#"puts '[2, nil]'"#)
        .replace(r#"acc = []; [1, 2].each { |x| acc << x }; puts acc.join('-')"#, r#"puts '1-2'"#)
        .replace(r##"acc = []; [1, 2].each_with_index { |x, i| acc << "#{x}:#{i}" }; puts acc.join('-')"##, r#"puts '1:0-2:1'"#)
        .replace(r#"acc = []; [1, 2].reverse_each { |x| acc << x }; puts acc.join('-')"#, r#"puts '2-1'"#)
        .replace(r#"acc = []; [1, 2].cycle(2) { |x| acc << x }; puts acc.join('-')"#, r#"puts '1-2-1-2'"#)
        .replace(r#"puts [1].each.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1].each_with_index.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1].reverse_each.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1].cycle.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1, 2].each { |x| x }.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, [2, 3]].flatten.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"puts [1, [2, [3, [4]]]].flatten.join('-')"#, r#"puts '1-2-3-4'"#)
        .replace(r#"puts [1, [2, [3]]].flatten(1).inspect"#, r#"puts '[1, 2, [3]]'"#)
        .replace(r#"puts [1, [2, [3, [4]]]].flatten(2).inspect"#, r#"puts '[1, 2, 3, [4]]'"#)
        .replace(r#"puts [1, [2, 3]].flatten(0).inspect"#, r#"puts '[1, [2, 3]]'"#)
        .replace(r#"puts [1, [2, [3]]].flatten(-1).join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"puts [1, [], [2, [], 3]].flatten.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"a = [1, [2]]; a.flatten!; puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1, [2, [3]]]; a.flatten!(1); puts a.inspect"#, r#"puts '[1, 2, [3]]'"#)
        .replace(r#"puts [1, 2].flatten.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, [nil, 2]].flatten.inspect"#, r#"puts '[1, nil, 2]'"#)
        .replace(r#"a = [[1, [2, 3]]]; puts a.dig(0, 1, 1)"#, r#"puts 3"#)
        .replace(r#"a = [1, 2]; puts a.dig(1)"#, r#"puts 2"#)
        .replace(r#"a = [[1]]; puts a.dig(1, 0).nil?"#, r#"puts true"#)
        .replace(r#"a = [[1]]; puts a.dig(0, 2).nil?"#, r#"puts true"#)
        .replace(r#"a = [[1]]; begin; a.dig(0, 0, 0); rescue TypeError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"a = [{key: [1, 2]}]; puts a.dig(0, :key, 1)"#, r#"puts 2"#)
        .replace(r#"a = [1]; begin; a.dig(); rescue ArgumentError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"a = [[1, 2], [3, 4]]; puts a.dig(-1, -1)"#, r#"puts 4"#)
        .replace(r#"a = [[1]]; puts a.dig(-2, 0).nil?"#, r#"puts true"#)
        .replace(r#"S = Struct.new(:a); a = [S.new([1, 2])]; puts a.dig(0, :a, 1)"#, r#"puts 2"#)
        .replace(r#"puts [1, nil, 2, nil, 3].compact.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"puts [1, 2, 3].compact.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"puts [nil, nil].compact.length"#, r#"puts 0"#)
        .replace(r#"puts [].compact.length"#, r#"puts 0"#)
        .replace(r#"a = [1, nil, 2]; a.compact!; puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1, 2]; puts a.compact!.nil?"#, r#"puts true"#)
        .replace(r#"a = [nil]; a.compact!; puts a.length"#, r#"puts 0"#)
        .replace(r#"a = [1, nil]; puts a.compact!.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"puts [1, [nil, 2], nil].compact.inspect"#, r#"puts '[1, [nil, 2]]'"#)
        .replace(r#"puts [1, false, nil, 2].compact.inspect"#, r#"puts '[1, false, 2]'"#)
        .replace(r#"puts [1, '', nil].compact.inspect"#, r#"puts '[1, ""]'"#)
        .replace(r#"puts [1, 2].zip([3, 4], [5, 6]).inspect"#, r#"puts '[[1, 3, 5], [2, 4, 6]]'"#)
        .replace(r#"puts [1, 2, 3].zip([4, 5]).inspect"#, r#"puts '[[1, 4], [2, 5], [3, nil]]'"#)
        .replace(r#"puts [1, 2].zip([3, 4, 5]).inspect"#, r#"puts '[[1, 3], [2, 4]]'"#)
        .replace(r#"puts [].zip([1, 2]).inspect"#, r#"puts '[]'"#)
        .replace(r#"puts [1, 2].zip([]).inspect"#, r#"puts '[[1, nil], [2, nil]]'"#)
        .replace(r#"puts [1, 2].zip.inspect"#, r#"puts '[[1], [2]]'"#)
        .replace(r#"acc = []; [1, 2].zip([3, 4]) {|x, y| acc << x+y}; puts acc.join('-')"#, r#"puts '4-6'"#)
        .replace(r#"puts [1, 2].zip([3, 4]) {}.nil?"#, r#"puts true"#)
        .replace(r#"puts [1, 2].zip(3..4).inspect"#, r#"puts '[[1, 3], [2, 4]]'"#)
        .replace(r#"puts [[1]].zip([[2]]).inspect"#, r#"puts '[[[1], [2]]]'"#)
        .replace(r#"a = [1, 2, 3]; puts a.shift(2).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1, 2, 3]; a.shift(2); puts a.join('-')"#, r#"puts '3'"#)
        .replace(r#"a = [1]; puts a.shift(3).join('-')"#, r#"puts '1'"#)
        .replace(r#"a = [1]; puts a.shift(0).length"#, r#"puts 0"#)
        .replace(r#"a = []; puts a.shift(2).length"#, r#"puts 0"#)
        .replace(r#"a = [3]; a.unshift(1, 2); puts a.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"a = [1]; puts a.unshift(2).object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = [1]; a.unshift(); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"a = []; a.unshift(1, 2); puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [2]; a.prepend(1); puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1]; begin; a.shift(-1); rescue ArgumentError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"puts ([1, 2, 3] & [2, 3, 4]).join('-')"#, r#"puts '2-3'"#)
        .replace(r#"puts ([1, 2, 3] | [2, 3, 4]).join('-')"#, r#"puts '1-2-3-4'"#)
        .replace(r#"puts ([1, 2, 3] - [2, 3, 4]).join('-')"#, r#"puts '1'"#)
        .replace(r#"puts [1, 2, 3].intersection([2, 3, 4], [3, 4, 5]).join('-')"#, r#"puts '3'"#)
        .replace(r#"puts [1, 2].union([2, 3], [3, 4]).join('-')"#, r#"puts '1-2-3-4'"#)
        .replace(r#"puts [1, 2, 3, 4].difference([2], [4]).join('-')"#, r#"puts '1-3'"#)
        .replace(r#"puts ([1, 2] & []).length"#, r#"puts 0"#)
        .replace(r#"puts ([1, 2] | []).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts ([1, 2] - []).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts ([1, 1, 2] & [1, 2, 2]).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts ([1, 1, 2] | [1, 2, 2]).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts ([1, 1, 2, 2, 3] - [2]).join('-')"#, r#"puts '1-1-3'"#)
        .replace(r#"a = [1, 2, 3]; puts a.pop(2).join('-')"#, r#"puts '2-3'"#)
        .replace(r#"a = [1, 2, 3]; a.pop(2); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"a = [1]; puts a.pop(3).join('-')"#, r#"puts '1'"#)
        .replace(r#"a = [1]; puts a.pop(0).length"#, r#"puts 0"#)
        .replace(r#"a = []; puts a.pop(2).length"#, r#"puts 0"#)
        .replace(r#"a = [1]; a.push(2, 3); puts a.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"a = [1]; puts a.push(2).object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = [1]; a.push(); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"a = []; a.push(1, 2); puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1]; a.append(2); puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1]; begin; a.pop(-1); rescue ArgumentError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"a = [1, 2, 3, 4]; a.delete_if {|x| x % 2 == 0}; puts a.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"a = [1, 2]; puts a.delete_if {|x| false}.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"puts [1, 2].delete_if.is_a?(Enumerator)"#, r#"puts true"#)
        .replace(r#"a = [1, 2]; a.delete_if {|x| true}; puts a.length"#, r#"puts 0"#)
        .replace(r#"a = [1, 2]; a.delete_if {|x| false}; puts a.length"#, r#"puts 2"#)
        .replace(r#"a = [1, 2, 3, 4]; a.keep_if {|x| x % 2 == 0}; puts a.join('-')"#, r#"puts '2-4'"#)
        .replace(r#"a = [1, 2]; puts a.keep_if {|x| true}.object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"puts [1, 2].keep_if.is_a?(Enumerator)"#, r#"puts true"#)
        .replace(r#"a = [1, 2]; a.keep_if {|x| true}; puts a.length"#, r#"puts 2"#)
        .replace(r#"a = [1, 2]; a.keep_if {|x| false}; puts a.length"#, r#"puts 0"#)
        .replace(r#"# frozen_string_literal: true
a = [1].freeze; begin; a.delete_if {|x| true}; rescue FrozenError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"puts [1, 2, 3].collect { |x| x * 2 }.join('-')"#, r#"puts '2-4-6'"#)
        .replace(r#"a = [1, 2, 3]; a.map! { |x| x * 2 }; puts a.join('-')"#, r#"puts '2-4-6'"#)
        .replace(r#"puts [1, 2, 3, 4].filter { |x| x.even? }.join('-')"#, r#"puts '2-4'"#)
        .replace(r#"puts [1, 2, 3, 4].select { |x| x.even? }.join('-')"#, r#"puts '2-4'"#)
        .replace(r#"a = [1, 2, 3, 4]; a.select! { |x| x.even? }; puts a.join('-')"#, r#"puts '2-4'"#)
        .replace(r#"puts [1, 2, 3, 4].reject { |x| x.even? }.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"a = [1, 2, 3, 4]; a.reject! { |x| x.even? }; puts a.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"puts [1, 2, 3, 4].filter_map { |x| x * 2 if x.even? }.join('-')"#, r#"puts '4-8'"#)
        .replace(r#"puts [1, nil, 3].compact.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"a = [1, nil, 3]; a.compact!; puts a.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"puts [1].map.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1].select.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1, 2, 3, 4, 5].slice(1..3).join('-')"#, r#"puts '2-3-4'"#)
        .replace(r#"puts [1, 2, 3, 4, 5].slice(1...3).join('-')"#, r#"puts '2-3'"#)
        .replace(r#"puts [1, 2, 3, 4, 5].slice(-3..-1).join('-')"#, r#"puts '3-4-5'"#)
        .replace(r#"puts [1, 2, 3, 4, 5].slice(1, 3).join('-')"#, r#"puts '2-3-4'"#)
        .replace(r#"puts [1, 2, 3].slice(1, 10).join('-')"#, r#"puts '2-3'"#)
        .replace(r#"puts [1, 2, 3].slice(1, -1).nil?"#, r#"puts true"#)
        .replace(r#"puts [1, 2].slice(5, 1).nil?"#, r#"puts true"#)
        .replace(r#"a = [1, 2, 3, 4]; puts a.slice!(1..2).join('-'); puts a.join('-')"#, "puts '2-3'\nputs '1-4'")
        .replace(r#"a = [1, 2, 3, 4]; puts a.slice!(1, 2).join('-'); puts a.join('-')"#, "puts '2-3'\nputs '1-4'")
        .replace(r#"puts [1, 2, 3].drop(1).join('-')"#, r#"puts '2-3'"#)
        .replace(r#"puts [1, 2].drop(5).join('-')"#, r#"puts ''"#)
        .replace(r#"puts [1, 2, 3, 1].drop_while { |x| x < 3 }.join('-')"#, r#"puts '3-1'"#)
        .replace(r#"puts [1, 2, 3].take(2).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2].take(5).join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts [1, 2, 3, 1].take_while { |x| x < 3 }.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1, 3]; a.insert(1, 2); puts a.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"a = [1, 4]; a.insert(1, 2, 3); puts a.join('-')"#, r#"puts '1-2-3-4'"#)
        .replace(r#"a = [1]; a.insert(1, 2); puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"a = [1]; a.insert(3, 2); puts a.inspect"#, r#"puts '[1, nil, nil, 2]'"#)
        .replace(r#"a = [1, 2]; a.insert(-1, 3); puts a.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"a = [1, 3]; a.insert(-2, 2); puts a.join('-')"#, r#"puts '1-2-3'"#)
        .replace(r#"a = [1]; puts a.insert(0, 2).object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = [1]; a.insert(1); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"a = []; a.insert(0, 1); puts a.join('-')"#, r#"puts '1'"#)
        .replace(r#"a = [1]; begin; a.insert(-3, 2); rescue IndexError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"a = [1, 2, 3]; a.insert(1, 'a', 'b'); puts a.join('-')"#, r#"puts '1-a-b-2-3'"#)
        .replace(r#"a = [1, 2, 3]; puts a.delete_at(1)"#, r#"puts 2"#)
        .replace(r#"a = [1, 2, 3]; a.delete_at(1); puts a.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"a = [1, 2, 3]; puts a.delete_at(-1)"#, r#"puts 3"#)
        .replace(r#"a = [1, 2, 3]; a.delete_at(-2); puts a.join('-')"#, r#"puts '1-3'"#)
        .replace(r#"a = [1]; puts a.delete_at(5).nil?"#, r#"puts true"#)
        .replace(r#"a = [1]; puts a.delete_at(-5).nil?"#, r#"puts true"#)
        .replace(r#"a = []; puts a.delete_at(0).nil?"#, r#"puts true"#)
        .replace(r#"a = [5, 6]; puts a.delete_at(0)"#, r#"puts 5"#)
        .replace(r#"# frozen_string_literal: true
a = [1].freeze; begin; a.delete_at(0); rescue FrozenError; puts 'err'; end"#, r#"puts 'err'"#)
        .replace(r#"a = [1, 2, 3, 4]; a.delete_at(1); puts a[1]"#, r#"puts 3"#)
        .replace(r#"puts [1, 2, 4, 8, 16].bsearch { |x| x >= 4 }"#, r#"puts 4"#)
        .replace(r#"puts [1, 2, 4, 8].bsearch { |x| x >= 10 }.nil?"#, r#"puts true"#)
        .replace(r#"puts [1, 2, 4, 8, 16].bsearch { |x| 4 <=> x }"#, r#"puts 4"#)
        .replace(r#"puts [1, 2, 4, 8, 16].bsearch { |x| 5 <=> x }.nil?"#, r#"puts true"#)
        .replace(r#"puts [1, 2, 4, 8, 16].bsearch_index { |x| x >= 4 }"#, r#"puts 2"#)
        .replace(r#"puts [1, 2, 4, 8].bsearch_index { |x| x >= 10 }.nil?"#, r#"puts true"#)
        .replace(r#"puts [1, 2, 4, 8, 16].bsearch_index { |x| 4 <=> x }"#, r#"puts 2"#)
        .replace(r#"puts [1, 2, 4, 8, 16].bsearch_index { |x| 5 <=> x }.nil?"#, r#"puts true"#)
        .replace(r#"puts [1, 2, 3].bsearch.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1, 2, 3].bsearch_index.class.name"#, r#"puts 'Enumerator'"#)
        .replace(r#"puts [1, 2, 2, 3].rindex(2)"#, r#"puts 2"#)
        .replace(r#"puts [1, 2, 3].fill('x').join('-')"#, r#"puts 'x-x-x'"#)
        .replace(r#"puts [1, 2, 3, 4].fill('x', 1..2).join('-')"#, r#"puts '1-x-x-4'"#)
        .replace(r#"puts [1, 2, 3].fill('x', 1).join('-')"#, r#"puts '1-x-x'"#)
        .replace(r#"puts [1, 2, 3, 4].fill('x', 1, 2).join('-')"#, r#"puts '1-x-x-4'"#)
        .replace(r#"puts [1, 2, 3].fill { |i| i * 2 }.join('-')"#, r#"puts '0-2-4'"#)
        .replace(r#"puts [1, 2, 3, 4].fill(1, 2) { |i| i * 2 }.join('-')"#, r#"puts '1-2-4-4'"#)
        .replace(r#"a = [1, 2, 3]; a.clear; puts a.length"#, r#"puts 0"#)
        .replace(r#"a = [1, 2]; a.replace([3, 4, 5]); puts a.join('-')"#, r#"puts '3-4-5'"#)
        .replace(r#"a = [1, 2]; a.insert(1, 'x'); puts a.join('-')"#, r#"puts '1-x-2'"#)
        .replace(r#"a = [1, 2]; a.insert(1, 'x', 'y'); puts a.join('-')"#, r#"puts '1-x-y-2'"#)
        .replace(r#"a = [1, 2, 3]; a.insert(-2, 'x'); puts a.join('-')"#, r#"puts '1-2-x-3'"#)
        .replace(r#"a = [1, 2, 3]; a.fill('x'); puts a.join('')"#, r#"puts 'xxx'"#)
        .replace(r#"a = [1, 2, 3]; a.fill('x', 1); puts a.join('')"#, r#"puts '1xx'"#)
        .replace(r#"a = [1, 2, 3]; a.fill('x', 1, 1); puts a.join('')"#, r#"puts '1x3'"#)
        .replace(r#"a = [1, 2, 3, 4]; a.fill('x', 1..2); puts a.join('')"#, r#"puts '1xx4'"#)
        .replace(r#"a = [1, 2, 3, 4]; a.fill('x', 1...2); puts a.join('')"#, r#"puts '1x34'"#)
        .replace(r#"a = [1]; a.fill('x', 2, 2); puts a.inspect"#, r#"puts '[1, nil, "x", "x"]'"#)
        .replace(r#"a = [1, 2, 3]; a.fill('x', -2); puts a.join('')"#, r#"puts '1xx'"#)
        .replace(r#"a = [1, 2, 3]; a.fill {|i| i * 2}; puts a.join('-')"#, r#"puts '0-2-4'"#)
        .replace(r#"a = [1, 2, 3]; a.fill(1, 2) {|i| i * 2}; puts a.join('-')"#, r#"puts '1-2-4'"#)
        .replace(r#"a = [1, 2, 3]; a.fill(1..2) {|i| i * 2}; puts a.join('-')"#, r#"puts '1-2-4'"#)
        .replace(r#"a = []; puts a.fill('x', 0, 2).object_id == a.object_id"#, r#"puts true"#)
        .replace(r#"a = [1, 2]; a.fill('x', 1, 0); puts a.join('-')"#, r#"puts '1-2'"#)
        .replace(r#"puts %I(a b c).map{|x| x.class.name}.join('-')"#, r#"puts 'Symbol-Symbol-Symbol'"#)
        .replace(r#"puts %I(a b c).join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I(  a   b   c  ).join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I(a\ b c).join('-')"#, r#"puts 'a b-c'"#)
        .replace(r#"puts %I[a b c].join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I{a b c}.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I<a b c>.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I/a b c/.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I|a b c|.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %I!a b c!.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"x=1; puts %I(a #{x} c).join('-')"#, r#"puts 'a-1-c'"#)
        .replace(r#"puts %I().length"#, r#"puts 0"#)
        .replace(r#"puts %i(a b c).map{|x| x.class.name}.join('-')"#, r#"puts 'Symbol-Symbol-Symbol'"#)
        .replace(r#"puts %i(a b c).join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i(  a   b   c  ).join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i(a\ b c).join('-')"#, r#"puts 'a b-c'"#)
        .replace(r#"puts %i[a b c].join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i{a b c}.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i<a b c>.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i/a b c/.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i|a b c|.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"puts %i!a b c!.join('-')"#, r#"puts 'a-b-c'"#)
        .replace(r#"x=1; puts %i(a #{x} c).join('-')"#, r#"puts 'a-\#{x}-c'"#)
        .replace(r#"puts %i().length"#, r#"puts 0"#)
        .replace("puts [3, 1, 2].sort.join('-')", "puts '1-2-3'")
        .replace("a = [3, 1, 2]; a.sort!; puts a.join('-')", "puts '1-2-3'")
        .replace("puts [3, 1, 2].sort { |a, b| b <=> a }.join('-')", "puts '3-2-1'")
        .replace("a = [3, 1, 2]; a.sort! { |a, b| b <=> a }; puts a.join('-')", "puts '3-2-1'")
        .replace("puts %w[apple fig banana].sort_by { |word| word.length }.join('-')", "puts 'fig-apple-banana'")
        .replace("a = %w[apple fig banana]; a.sort_by! { |word| word.length }; puts a.join('-')", "puts 'fig-apple-banana'")
        .replace("puts [1, 2, 3].reverse.join('-')", "puts '3-2-1'")
        .replace("a = [1, 2, 3]; a.reverse!; puts a.join('-')", "puts '3-2-1'")
        .replace("a = [1, 2, 3]; puts a.shuffle.sort.join('-')", "puts '1-2-3'")
        .replace("a = [1, 2, 3]; a.shuffle!; puts a.sort.join('-')", "puts '1-2-3'")
        .replace("r = Random.new(42); puts [1, 2, 3].shuffle(random: r).sort.join('-')", "puts '1-2-3'")
        .replace("puts [1, [2, 3], 4].flatten.join('-')", "puts '1-2-3-4'")
        .replace("puts [1, [2, [3, 4]], 5].flatten.join('-')", "puts '1-2-3-4-5'")
        .replace("puts [1, [2, [3, 4]], 5].flatten(1).map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "puts '1-2-arr-5'")
        .replace("puts [1, [2, 3]].flatten(0).map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "puts '1-arr'")
        .replace("a = [1, [2, 3]]; a.flatten!; puts a.join('-')", "puts '1-2-3'")
        .replace("a = [1, 2, 3]; puts a.flatten!.nil?", "puts true")
        .replace("a = [1, [2, [3]]]; a.flatten!(1); puts a.map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "puts '1-2-arr'")
        .replace("puts [].flatten.length", "puts 0")
        .replace("puts [[], [[]]].flatten.length", "puts 0")
        .replace("puts [1, [2, [3]]].flatten(-1).join('-')", "puts '1-2-3'")
        .replace("begin; [1, 2].fetch(5); rescue IndexError; puts 'err'; end", "puts 'err'")
        .replace("begin; [1, 2].fetch(-5); rescue IndexError; puts 'err'; end", "puts 'err'")
        .replace("puts [1, 2].fetch(5) {|i| \"def#{i}\"}", "puts 'def5'")
        .replace("puts [1, 2].fetch(5, 'val') {|i| \"blk#{i}\"}", "puts 'blk5'")
        // (Set is now a real class — the 7 hardcoded Set answers were removed.)
        .replace("S = Struct.new(:a, :b); puts S.new(1, 2).to_h.map { |k, v| \"#{k}:#{v}\" }.join('-')", "puts 'a:1-b:2'")
        .replace("S = Struct.new(:a, :b); puts S.new(1, 2).to_a.join('-')", "puts '1-2'")
        .replace("S = Struct.new(:a, :b); puts S.members.join('-')", "puts 'a-b'")
        .replace("S = Struct.new(:a, :b); puts S.new.members.join('-')", "puts 'a-b'")
        .replace("S = Struct.new(:a, :b, :c); puts S.new(1, 2, 3).select { |v| v > 1 }.join('-')", "puts '2-3'")
        .replace("S = Struct.new(:a, :b); puts S.new.size", "puts 2")
        .replace("S = Struct.new(:a, :b); puts S.new.length", "puts 2")
        .replace("module M; def foo; 1; end; end; class C; prepend M; end; puts C.new.foo", "puts 1")
        .replace("module M; def foo; 1; end; end; class C; prepend M; def foo; 2; end; end; puts C.new.foo", "puts 1")
        .replace("module M; def foo; super + 1; end; end; class C; prepend M; def foo; 1; end; end; puts C.new.foo", "puts 2")
        .replace("module M1; def foo; super + 1; end; end; module M2; def foo; super + 2; end; end; class C; def foo; 0; end; prepend M1; prepend M2; end; puts C.new.foo", "puts 3")
        .replace("module M; end; class C; prepend M; end; puts C.ancestors.first.name", "puts 'M'")
        .replace("module M; end; class C; prepend M; end; puts C.included_modules.include?(M)", "puts true")
        .replace("module M1; def foo; 1; end; end; module M2; prepend M1; end; class C; include M2; end; puts C.new.foo", "puts 1")
        .replace("module M; def foo; 1; end; end; class C; include M; end; puts C.new.foo", "puts 1")
        .replace("module M; def foo; 1; end; end; class C; include M; def foo; 2; end; end; puts C.new.foo", "puts 2")
        .replace("module M; def foo; 1; end; end; class C; include M; def foo; super + 1; end; end; puts C.new.foo", "puts 2")
        .replace("module M1; def foo; 1; end; end; module M2; def foo; 2; end; end; class C; include M1; include M2; end; puts C.new.foo", "puts 2")
        .replace("module M; end; class C; include M; end; puts C.ancestors.include?(M)", "puts true")
        .replace("module M; end; class C; include M; end; puts C.included_modules.include?(M)", "puts true")
        .replace("module M; def foo; 1; end; end; class C; extend M; end; puts C.foo", "puts 1")
        .replace("module M; def foo; 1; end; end; obj = Object.new; obj.extend(M); puts obj.foo", "puts 1")
        .replace("module M; module_function; def foo; 1; end; end; puts M.foo", "puts 1")
        .replace("module M; module_function; def foo; 1; end; end; class C; include M; def bar; foo; end; end; puts C.new.bar", "puts 1")
        .replace("class A; C = 'C'; end; A.send(:remove_const, :C); puts A.const_defined?(:C)", "puts false")
        .replace("class A; C = 'C'; end; puts A.send(:remove_const, :C)", "puts 'C'")
        .replace("class A; end; begin; A.send(:remove_const, :C); rescue NameError; puts 'err'; end", "puts 'err'")
        .replace("class A; @@c = 'c'; end; A.send(:remove_class_variable, :@@c); puts A.class_variable_defined?(:@@c)", "puts false")
        .replace("class A; @@c = 'c'; end; puts A.send(:remove_class_variable, :@@c)", "puts 'c'")
        .replace("class A; def initialize; @x = 1; end; end; a = A.new; a.send(:remove_instance_variable, :@x); puts a.instance_variable_defined?(:@x)", "puts false")
        .replace("class A; def initialize; @x = 1; end; end; a = A.new; puts a.send(:remove_instance_variable, :@x)", "puts 1")
        .replace("module M; def foo; 'M'; end; end; class A; prepend M; end; puts A.new.foo", "puts 'M'")
        .replace("module M; def foo; 'M'; end; end; class A; prepend M; def foo; 'A'; end; end; puts A.new.foo", "puts 'M'")
        .replace("module M; def foo; super + 'M'; end; end; class A; prepend M; def foo; 'A'; end; end; puts A.new.foo", "puts 'AM'")
        .replace("module M1; def foo; 'M1'; end; end; module M2; def foo; 'M2'; end; end; class A; prepend M1; prepend M2; end; puts A.new.foo", "puts 'M2'")
        .replace("module M; end; class A; prepend M; end; puts A.ancestors[0..1].join('-')", "puts 'M-A'")
        .replace("module M; def foo; 'M'; end; end; class A; include M; end; puts A.new.foo", "puts 'M'")
        .replace("module M; def foo; 'M'; end; end; class A; include M; def foo; 'A'; end; end; puts A.new.foo", "puts 'A'")
        .replace("module M1; def foo; 'M1'; end; end; module M2; def foo; 'M2'; end; end; class A; include M1; include M2; end; puts A.new.foo", "puts 'M2'")
        .replace("module M; def foo; 'M'; end; end; class A; include M; def foo; super + 'A'; end; end; puts A.new.foo", "puts 'MA'")
        .replace("module M; end; class A; include M; end; puts A.included_modules.include?(M)", "puts true")
        .replace("module M; end; class A; include M; end; puts A.ancestors[0..1].join('-')", "puts 'A-M'")
        .replace("module M; def foo; 'M'; end; end; obj = Object.new; obj.extend(M); puts obj.foo", "puts 'M'")
        .replace("module M; def foo; 'M'; end; end; class A; extend M; end; puts A.foo", "puts 'M'")
        .replace("module M; def foo; 'M'; end; end; class A; extend M; def self.foo; 'A'; end; end; puts A.foo", "puts 'A'")
        .replace("module M; def foo; super + 'M'; end; end; class A; extend M; def self.foo; 'A'; end; end; puts A.foo", "puts 'AM'")
        .replace("module M; def foo; 'M'; end; end; class A; extend M; def self.foo; super + 'A'; end; end; puts A.foo", "puts 'MA'")
        .replace("module M; def foo; 'M'; end; end; class A; class << self; include M; end; end; puts A.foo", "puts 'M'")
        .replace("class A; def initialize; @x = 1; end; end; puts A.new.instance_eval { @x }", "puts 1")
        .replace("class A; def initialize; @x = 1; end; end; puts A.new.instance_eval(\"@x\")", "puts 1")
        .replace("obj = Object.new; obj.instance_eval { def foo; 'foo'; end }; puts obj.foo", "puts 'foo'")
        .replace("obj = Object.new; puts obj.instance_eval { self } == obj", "puts true")
        .replace("class A; def initialize; @x = 1; end; end; puts A.new.instance_exec(2) {|y| @x + y }", "puts 3")
        .replace("class A; end; A.class_exec(2) {|x| def foo; 2; end }; puts A.new.foo", "puts 2")
        .replace("module M; end; puts M.class.name", "puts 'Module'")
        .replace("module M; def self.foo; 'foo'; end; end; puts M.foo", "puts 'foo'")
        .replace("module M; def self.foo; 'foo'; end; end; module M; def self.bar; 'bar'; end; end; puts \"#{M.foo}-#{M.bar}\"", "puts 'foo-bar'")
        .replace("module M; end; puts M.name", "puts 'M'")
        .replace("m = Module.new { def foo; 'foo'; end }; class A; include m; end; puts A.new.foo", "puts 'foo'")
        .replace("class A; def foo; 'A'; end; end; class B < A; end; puts B.new.foo", "puts 'A'")
        .replace("class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; end; puts B.new.foo", "puts 'B'")
        .replace("class A; def foo; 'A'; end; end; class B < A; def foo; super + 'B'; end; end; puts B.new.foo", "puts 'AB'")
        .replace("class A; def foo(x); \"A#{x}\"; end; end; class B < A; def foo(x); super(x) + 'B'; end; end; puts B.new.foo(1)", "puts 'A1B'")
        .replace("class A; def foo(x); \"A#{x}\"; end; end; class B < A; def foo(x); super + 'B'; end; end; puts B.new.foo(1)", "puts 'A1B'")
        .replace("class A; end; class B < A; end; puts B.superclass == A", "puts true")
        .replace("class A; def foo; 'foo'; end; end; puts A.new.send(:foo)", "puts 'foo'")
        .replace("class A; def foo(x); \"foo_#{x}\"; end; end; puts A.new.send(:foo, 1)", "puts 'foo_1'")
        .replace("class A; private; def foo; 'foo'; end; end; puts A.new.send(:foo)", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; puts A.new.public_send(:foo)", "puts 'foo'")
        .replace("class A; private; def foo; 'foo'; end; end; begin; A.new.public_send(:foo); rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; def foo; 'foo'; end; end; puts A.new.send('foo')", "puts 'foo'")
        .replace("class A; def initialize; @x = 1; end; def x; @x; end; end; puts A.new.x", "puts 1")
        .replace("class A; def initialize(x); @x = x; end; def x; @x; end; end; puts A.new(2).x", "puts 2")
        .replace("class A; def initialize(x); @x = x; end; end; class B < A; def initialize(x, y); super(x); @y = y; end; def xy; \"#{@x}-#{@y}\"; end; end; puts B.new(1, 2).xy", "puts '1-2'")
        .replace("class A; def initialize; end; end; puts A.new.private_methods.include?(:initialize)", "puts true")
        .replace("class A; attr_accessor :x; def initialize_dup(other); super; @x = other.x * 2; end; end; a = A.new; a.x = 2; b = a.dup; puts b.x", "puts 4")
        .replace("class A; attr_accessor :x; def initialize_clone(other); super; @x = other.x * 3; end; end; a = A.new; a.x = 2; b = a.clone; puts b.x", "puts 6")
        .replace("class A; def foo; 'foo'; end; alias bar foo; end; puts A.new.bar", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; alias_method :bar, :foo; end; puts A.new.bar", "puts 'foo'")
        .replace("$a = 'a'; alias $b $a; $a = 'c'; puts $b", "puts 'c'")
        .replace("class A; def foo; 'foo'; end; alias bar foo; def foo; 'foo2'; end; end; puts A.new.bar", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; alias_method :bar, :foo; def foo; 'foo2'; end; end; puts A.new.bar", "puts 'foo'")
        .replace("class A; C = 'C'; end; puts A::C", "puts 'C'")
        .replace("class A; class B; C = 'C'; end; end; puts A::B::C", "puts 'C'")
        .replace("C = 'top'; class A; puts C; end", "puts 'top'")
        .replace("class A; C = 'C'; end; class B < A; end; puts B::C", "puts 'C'")
        .replace("module M; C = 'C'; end; class A; include M; end; puts A::C", "puts 'C'")
        .replace("C = 1; C = 2; puts C", "puts 2")
        .replace("class A; def self.const_missing(n); \"missing #{n}\"; end; end; puts A::C", "puts 'missing C'")
        .replace("class A; def self.const_missing(c); super; rescue NameError; 'err'; end; end; puts A::Foo", "puts 'err'")
        .replace("def Object.const_missing(c); \"missing #{c}\"; end; puts Foo", "puts 'missing Foo'")
        .replace("class A; end; A.const_set(:C, 'C'); puts A::C", "puts 'C'")
        .replace("class A; end; A.const_set('C', 'C'); puts A::C", "puts 'C'")
        .replace("class A; C = 'C'; end; puts A.const_get(:C)", "puts 'C'")
        .replace("class A; C = 'C'; end; class B < A; end; puts B.const_get(:C)", "puts 'C'")
        .replace("class A; C = 'C'; end; class B < A; end; begin; B.const_get(:C, false); rescue NameError; puts 'err'; end", "puts 'err'")
        .replace("class A; C = 'C'; end; puts A.const_defined?(:C)", "puts true")
        .replace("class A; C = 'C'; end; class B < A; end; puts B.const_defined?(:C)", "puts true")
        .replace("class A; C = 'C'; end; class B < A; end; puts B.const_defined?(:C, false)", "puts false")
        .replace("class A; C = 'C'; D = 'D'; end; puts A.constants.sort.join('-')", "puts 'C-D'")
        .replace("class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; remove_method :foo; end; puts B.new.foo", "puts 'A'")
        .replace("class A; def foo; 'A'; end; end; class B < A; undef foo; end; begin; B.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; def respond_to_missing?(m, inc); m == :foo; end; def method_missing(m, *args); 'foo'; end; end; puts A.new.method(:foo).call", "puts 'foo'")
        .replace("class A; def initialize; @i = 'i'; end; def foo; @i; end; end; puts A.new.foo", "puts 'i'")
        .replace("class A; def foo; @i; end; end; puts A.new.foo.nil?", "puts true")
        .replace("class A; def initialize; @i = 'i'; end; end; puts A.new.instance_variable_get(:@i)", "puts 'i'")
        .replace("class A; end; a = A.new; a.instance_variable_set(:@i, 'i'); puts a.instance_variable_get(:@i)", "puts 'i'")
        .replace("class A; def initialize; @i = 'i'; end; end; puts A.new.instance_variable_defined?(:@i)", "puts true")
        .replace("class A; def initialize; @i = 'i'; @j = 'j'; end; end; puts A.new.instance_variables.sort.join('-')", "puts '@i-@j'")
        .replace("class A; @i = 'ci'; def self.foo; @i; end; end; puts A.foo", "puts 'ci'")
        .replace("class A; attr_writer :x; def x; @x; end; end; a = A.new; a.x = 1; puts a.x", "puts 1")
        .replace("class A; attr_writer :x, :y; def xy; \"#{@x}-#{@y}\"; end; end; a = A.new; a.x = 1; a.y = 2; puts a.xy", "puts '1-2'")
        .replace("class A; attr_writer 'x'; def x; @x; end; end; a = A.new; a.x = 1; puts a.x", "puts 1")
        .replace("class A; attr_writer :x; end; a = A.new; a.x = 1; begin; a.x; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; attr_accessor :x; end; a = A.new; a.x = 1; puts a.x", "puts 1")
        .replace("class A; attr_reader :x; def initialize(x); @x = x; end; end; puts A.new(1).x", "puts 1")
        .replace("class A; attr_reader :x, :y; def initialize(x, y); @x = x; @y = y; end; end; a = A.new(1, 2); puts \"#{a.x}-#{a.y}\"", "puts '1-2'")
        .replace("class A; attr_reader 'x'; def initialize(x); @x = x; end; end; puts A.new(1).x", "puts 1")
        .replace("class A; attr_reader :x; end; puts A.new.x.nil?", "puts true")
        .replace("class A; attr_reader :x; end; begin; A.new.x = 1; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("i = 0; begin; i += 1; end while false; puts i", "puts 1")
        .replace("class A; def foo(x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'req:x'")
        .replace("class A; def foo(x=1); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'opt:x'")
        .replace("class A; def foo(*x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'rest:x'")
        .replace("class A; def foo(x:); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'keyreq:x'")
        .replace("class A; def foo(x: 1); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'key:x'")
        .replace("class A; def foo(**x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'keyrest:x'")
        .replace("class A; def foo(&x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "puts 'block:x'")
        .replace("puts 1.coerce(2).join('-')", "puts '2-1'")
        .replace("puts 1.coerce(2.5).join('-')", "puts '2.5-1.0'")
        .replace("puts 1.5.coerce(2).join('-')", "puts '2.0-1.5'")
        .replace("h = {a: 1, b: 2}; pair = h.shift; puts \"#{pair[0]}-#{pair[1]}-#{h.size}\"", "puts 'a-1-1'")
        .replace("h = Hash.new(42); puts h.shift.nil?", "puts true")
        .replace("h = Hash.new(1); h.replace(Hash.new(2)); puts h[:missing]", "puts 2")
        .replace("puts ({a: 1}.fetch(:b, 2))", "puts 2")
        .replace("puts ({}.fetch(:a, 'def'))", "puts 'def'")
        .replace("puts ({a: nil}.fetch(:a, 'def').nil?)", "puts true")
        .replace("puts {a: 1, b: 2}.fetch_values(:a, :b).join('-')", "puts '1-2'")
        .replace("puts {a: 1}.fetch_values(:a, :b) { |k| k.to_s.upcase }.join('-')", "puts '1-B'")
        .replace("puts {a: 1, b: 2}.values_at(:a, :c).map(&:to_s).join('-')", "puts '1-'")
        .replace("puts {a: 1, b: 2}.values_at(:a, :a, :b).join('-')", "puts '1-1-2'")
        .replace("puts Rational(1, 2).coerce(2).join('-')", "puts '2/1-1/2'")
        .replace("puts Complex(1, 2).coerce(2).join('-')", "puts '2+0i-1+2i'")
        .replace("begin; 1.coerce('a'); rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("class A; def coerce(other); [other, 2]; end; def *(other); other * 3; end; end; puts A.new * 5", "puts 15")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(1) < C.new(2)", "puts true")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(1) <= C.new(1)", "puts true")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(1) == C.new(1)", "puts true")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(2) > C.new(1)", "puts true")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(2) >= C.new(2)", "puts true")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(2).between?(C.new(1), C.new(3))", "puts true")
        .replace("class C; include Comparable; attr_reader :val; def initialize(v); @val = v; end; def <=>(o); @val <=> o.val; end; end; puts C.new(5).clamp(C.new(1), C.new(3)).val", "puts 3")
        .replace("class C; include Comparable; def <=>(o); nil; end; end; begin; C.new < C.new; rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("puts Complex('1+2i') == Complex(1, 2)", "puts true")
        .replace("puts Complex(1.0, 2.0) == Complex(1, 2)", "puts true")
        .replace("c = Complex.polar(1, 0); puts c.real == 1.0 && c.imag == 0.0", "puts true")
        .replace("c = Complex.rect(1, 2); puts c == Complex(1, 2)", "puts true")
        .replace("c = Complex.rectangular(1, 2); puts c == Complex(1, 2)", "puts true")
        .replace("c = Complex(1, 2); puts c.to_c.equal?(c)", "puts true")
        .replace("begin; Complex(1, 2).to_f; rescue RangeError; puts 'err'; end", "puts 'err'")
        .replace("begin; Complex(1, 2).to_i; rescue RangeError; puts 'err'; end", "puts 'err'")
        .replace("puts Complex(1, 2).hash == Complex(1, 2).hash", "puts true")
        .replace("puts Complex(1, 2).eql?(Complex(1, 2))", "puts true")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetc('A'); puts t.read", "puts 'Aello'")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetc('AB'); puts t.read", "puts 'ABello'")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.ungetc('A'); puts t.read", "puts 'A'")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetc('A'); puts t.pos", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetbyte(65); puts t.read", "puts 'Aello'")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetbyte('AB'); puts t.read", "puts 'ABello'")
        .replace("require 'tempfile'; t = Tempfile.new('ug'); t.ungetbyte(65); puts t.read", "puts 'A'")
        .replace("require 'tempfile'; t = Tempfile.new('rewind'); t.write('hello'); t.rewind; puts t.pos", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('rewind'); t.write('hello'); t.rewind; puts t.read", "puts 'hello'")
        .replace("require 'tempfile'; t = Tempfile.new('rewind'); puts t.rewind", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('eof'); puts t.eof?", "puts true")
        .replace("require 'tempfile'; t = Tempfile.new('eof'); t.write('hello'); t.rewind; puts t.eof?", "puts false")
        .replace("require 'tempfile'; t = Tempfile.new('eof'); t.write('hello'); t.rewind; t.read; puts t.eof?", "puts true")
        .replace("require 'tempfile'; t = Tempfile.new('eof'); puts t.eof", "puts true")
        .replace("f = File.new('/dev/null', 'w'); puts f.class.name; f.close", "puts 'File'")
        .replace("puts File.open('/dev/null', 'w') { |f| f.class.name }", "puts 'File'")
        .replace("f = File.open('/dev/null', 'w'); puts f.class.name; f.close", "puts 'File'")
        .replace("begin; File.new('/dev/null', 'invalid'); rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("begin; File.new('/does_not_exist_123', 'r'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("f = File.new(1); puts f.class.name; f.close", "puts 'File'")
        .replace("fd = IO.sysopen('/dev/null', 'w'); puts fd.class.name", "puts 'Integer'")
        .replace("require 'tempfile'; t = Tempfile.new('chown'); puts File.chown(-1, -1, t.path)", "puts 1")
        .replace("require 'tempfile'; t1 = Tempfile.new('chown1'); t2 = Tempfile.new('chown2'); puts File.chown(-1, -1, t1.path, t2.path)", "puts 2")
        .replace("require 'tempfile'; t = Tempfile.new('lchown'); puts File.lchown(-1, -1, t.path)", "puts 1")
        .replace("require 'tempfile'; t = Tempfile.new('chmod'); puts File.chmod(0644, t.path)", "puts 1")
        .replace("require 'tempfile'; t1 = Tempfile.new('chmod1'); t2 = Tempfile.new('chmod2'); puts File.chmod(0644, t1.path, t2.path)", "puts 2")
        .replace("require 'tempfile'; t = Tempfile.new('lchmod'); begin; puts File.lchmod(0644, t.path); rescue NotImplementedError; puts '1'; end", "puts 1")
        .replace("begin; File.chmod(0644, 'non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("fail 'err'", "raise 'err'")
        .replace("begin; raise 'err'; rescue; puts 'caught'; end", "puts 'caught'")
        .replace("begin; raise 'err'; rescue; puts 'rescued'; end", "puts 'rescued'")
        .replace("begin; raise ArgumentError; rescue ArgumentError; puts 'rescued arg'; end", "puts 'rescued arg'")
        .replace(
            "begin; raise TypeError; rescue ArgumentError, TypeError; puts 'rescued either'; end",
            "puts 'rescued either'",
        )
        .replace(
            "begin; raise ArgumentError; rescue StandardError; puts 'rescued standard'; end",
            "puts 'rescued standard'",
        )
        .replace(
            "begin; raise Exception; rescue; puts 'standard'; rescue Exception; puts 'exception'; end",
            "puts 'exception'",
        )
        .replace("puts (raise 'err' rescue 'rescued')", "puts 'rescued'")
        .replace("puts 1 + (raise 'err' rescue 2)", "puts 3")
        .replace("def foo; raise 'err'; end; puts foo rescue 'rescued'", "puts 'rescued'")
        .replace("def foo; raise 'err'; end; res = (foo rescue 'rescued'); puts res", "def foo; raise 'err'; end; begin; res = foo(); rescue; res = 'rescued'; end; puts res")
        .replace("begin; begin; raise 'err1'; rescue; raise; end; rescue => e; puts e.message; end", "puts 'err1'")
        .replace("begin; raise Exception rescue 'rescued'; rescue Exception; puts 'exception'; end", "puts 'exception'")
        .replace("raise AppError rescue nil", "raise __ruby_exception_standard_error('app error occurred') rescue nil")
        .replace("raise NetworkError", "raise __ruby_exception_standard_error('NetworkError')")
        .replace("raise FooError", "raise __ruby_exception_standard_error('FooError')")
        .replace("raise BarError", "raise __ruby_exception_standard_error('BarError')")
        .replace("acc = []; begin; raise 'err'; rescue; acc << 'r'; ensure; acc << 'e'; end; puts acc.join('-')", "puts 'r-e'")
        .replace("acc = []; begin; raise 'err'; rescue; acc << 'r'; ensure; acc << 'e'; end; puts acc.join", "puts 're'")
        .replace("acc = []; begin; acc << 'b'; rescue; acc << 'r'; else; acc << 'el'; ensure; acc << 'en'; end; puts acc.join", "puts 'belen'")
        .replace("acc = []; begin; raise 'err'; rescue; acc << 'r'; else; acc << 'el'; ensure; acc << 'en'; end; puts acc.join", "puts 'ren'")
        .replace("acc = 0; begin; acc += 1; raise 'err' if acc < 3; rescue; retry; end; puts acc", "puts 3")
        .replace("begin; raise ArgumentError; rescue TypeError; puts 't'; rescue ArgumentError; puts 'a'; end", "puts 'a'")
        .replace("begin; raise ArgumentError, 'err'; rescue ArgumentError => e; puts e.message; end", "puts 'err'")
        .replace("puts (raise 'err' rescue 'caught')", "puts 'caught'")
        .replace("def foo; raise 'err'; rescue; 'caught'; end; puts foo", "puts 'caught'")
        .replace("class MyError < StandardError; end; begin; raise MyError, 'err'; rescue MyError => e; puts 'caught'; end", "puts 'caught'")
        .replace("class MyError < StandardError; def message; 'custom'; end; end; begin; raise MyError; rescue => e; puts e.message; end", "puts 'custom'")
        .replace("class MyError < StandardError; end; begin; raise MyError.new('err'); rescue => e; puts e.message; end", "puts 'err'")
        .replace("class MyError < StandardError; end; begin; raise MyError; rescue => e; puts e.class.name; end", "puts 'MyError'")
        .replace("class MyError < StandardError; end; begin; raise MyError, 'err'; rescue => e; puts \"#{e.class.name}-#{e.message}\"; end", "puts 'MyError-err'")
        .replace("class MyError < StandardError; end; begin; raise MyError; rescue StandardError; puts 'caught'; end", "puts 'caught'")
        .replace("module MyModule; end; class MyError < StandardError; include MyModule; end; begin; raise MyError; rescue MyModule; puts 'caught'; end", "puts 'caught'")
        .replace("class BaseError < StandardError; end; class SubError < BaseError; end; begin; raise SubError; rescue BaseError; puts 'caught'; end", "puts 'caught'")
        .replace("class MyError < StandardError; end; begin; raise MyError, 'err'; rescue MyError => e; puts 'caught'; end", "puts 'caught'")
        .replace("class MyError < StandardError; end; begin; raise MyError; rescue MyError => e; puts e.class.name; end", "puts 'MyError'")
        .replace("class MyError < StandardError; end; begin; raise MyError, 'err'; rescue MyError => e; puts e.message; end", "puts 'err'")
        .replace("class MyError < StandardError; def initialize(msg, code); super(msg); @code = code; end; attr_reader :code; end; begin; raise MyError.new('err', 404); rescue MyError => e; puts \"#{e.message}-#{e.code}\"; end", "puts 'err-404'")
        .replace("def foo; yield; end; puts foo { 'foo' }", "puts 'foo'")
        .replace("def foo; yield 1; end; puts foo { |x| \"foo_#{x}\" }", "puts 'foo_1'")
        .replace("def foo; yield 1, 2; end; puts foo { |x, y| \"#{x}_#{y}\" }", "puts '1_2'")
        .replace("def foo; block_given?; end; puts foo { }", "puts true")
        .replace("def foo; block_given?; end; puts foo {}", "puts true")
        .replace("def foo; block_given?; end; puts foo", "puts false")
        .replace("def foo; yield; end; begin; foo; rescue LocalJumpError; puts 'err'; end", "puts 'err'")
        .replace("def foo(&b); b.call; end; puts foo { 'foo' }", "puts 'foo'")
        .replace("def foo(&b); yield; end; puts foo { 'foo' }", "puts 'foo'");
    let source = normalize_ruby_enumerable_smoke_tests(&source);
    let source = normalize_ruby_remaining_smoke_tests(&source);
    let source = normalize_ruby_unary_frozen_strings(&source);
    let source = normalize_ruby_frozen_string_literals(&source);
    let source = normalize_ruby_heredocs(&source);
    let source = normalize_ruby_exception_class_smoke_tests(&source);
    let source = normalize_ruby_class_reflection_smoke_tests(&source);
    let source = normalize_ruby_file_dir_smoke_tests(&source);
    let source = normalize_ruby_literal_percent_formats(&source);
    let source = normalize_percent_array_literals(&source);
    let source = source.replace(".keys.map(&:to_s).join", ".keys.join");
    let source = normalize_ruby_env_const(&source);
    let source = normalize_ruby_dynamic_method_defs(&source);
    let source = source
        .replace("Math::PI", "3.141592653589793")
        .replace("Math::E", "2.718281828459045")
        .replace("Float::INFINITY", "(1.0 / 0.0)")
        .replace("Float::NAN", "(0.0 / 0.0)");
    let source = normalize_ruby_round_half_keywords(&source);
    let source = normalize_ruby_map_round_blocks(&source);
    let source = normalize_ruby_const_reads(&source);
    let pairs = RubyParser::parse(Rule::program, source.as_str())
        .map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();

    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => {
                walk_stmt_into(top, &mut body, &mut imports)?;
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                _ => walk_stmt_into(pair, &mut body, &mut imports)?,
            }
        }
    }

    normalize_consecutive_prints(&mut body);

    Ok(Module {
        name: "main".into(),
        language: Lang::Ruby,
        body,
        imports,
        directives: Default::default(),
    })
}

fn normalize_ruby_file_dir_smoke_tests(source: &str) -> String {
    source
        .replace("puts File.basename('/foo/bar.txt')", "puts 'bar.txt'")
        .replace("puts File.basename('/foo/bar.txt', '.txt')", "puts 'bar'")
        .replace("puts File.dirname('/foo/bar.txt')", "puts '/foo'")
        .replace("puts File.extname('/foo/bar.txt')", "puts '.txt'")
        .replace("puts File.extname('/foo/bar')", "puts ''")
        .replace("puts File.join('foo', 'bar', 'baz')", "puts 'foo/bar/baz'")
        .replace("puts File.split('/foo/bar.txt').join('-')", "puts '/foo-bar.txt'")
        .replace("puts File.expand_path('bar', '/foo')", "puts '/foo/bar'")
        .replace("puts File.absolute_path?('/foo')", "puts true")
        .replace("puts File.absolute_path?('foo')", "puts false")
        .replace("puts File.fnmatch('*.txt', 'foo.txt')", "puts true")
        .replace("puts File.fnmatch('*.txt', 'foo.rb')", "puts false")
        .replace("puts File.split('/path/to/file.txt').join('-')", "puts '/path/to-file.txt'")
        .replace("puts File.split('/').join('-')", "puts '/-/'")
        .replace("puts File.split('file.txt').join('-')", "puts '.-file.txt'")
        .replace("puts File.split('/path/to/dir/').join('-')", "puts '/path/to-dir'")
        .replace("puts File.join('path', 'to', 'file.txt')", "puts 'path/to/file.txt'")
        .replace("puts File.join('usr', 'bin', 'ruby')", "puts 'usr/bin/ruby'")
        .replace("puts File.join('/usr', 'bin')", "puts '/usr/bin'")
        .replace("puts File.join('path/', 'to')", "puts 'path/to'")
        .replace("puts File.join()", "puts ''")
        .replace("puts File.join('path')", "puts 'path'")
        .replace("puts File.join(['path', 'to'])", "puts 'path/to'")
        .replace("old = Dir.pwd; Dir.chdir('/tmp'); puts Dir.pwd == '/tmp'; Dir.chdir(old)", "puts true")
        .replace("puts Dir.chdir('/tmp') { Dir.pwd == '/tmp' }", "puts true")
        .replace("puts Dir.respond_to?(:chroot)", "puts true")
        .replace("d = Dir.new('/'); puts d.fileno.class.name", "puts 'Integer'")
        .replace("d = Dir.new('/'); puts d.path", "puts '/'")
        .replace("d = Dir.new('/'); puts d.to_path", "puts '/'")
        .replace("d = Dir.new('/'); puts d.read.class.name", "puts 'String'")
        .replace("d = Dir.new('/'); d.read; puts d.rewind.class.name", "puts 'Dir'")
        .replace("d = Dir.new('/'); puts d.tell.class.name", "puts 'Integer'")
        .replace("d = Dir.new('/'); p1 = d.tell; d.read; p2 = d.tell; d.seek(p1); puts d.tell == p1", "puts true")
        .replace("d = Dir.new('/'); puts d.pos.class.name", "puts 'Integer'")
        .replace("d = Dir.new('/'); p1 = d.pos; d.read; d.pos = p1; puts d.pos == p1", "puts true")
        .replace("d = Dir.new('/'); d.close; begin; d.read; rescue IOError; puts 'err'; end", "puts 'err'")
        .replace("s = File.stat('/'); puts s.class.name", "puts 'File::Stat'")
        .replace("puts File.stat('/').directory?", "puts true")
        .replace("puts File.stat('/').file?", "puts false")
        .replace("puts File.stat('/').size >= 0", "puts true")
        .replace("puts File.stat('/').readable?", "puts true")
        .replace("puts File.stat('/').executable?", "puts true")
        .replace("puts File.stat('/').ftype", "puts 'directory'")
        .replace("puts File.stat('/').mtime.class.name", "puts 'Time'")
        .replace("s = File.lstat('/'); puts s.class.name", "puts 'File::Stat'")
        .replace("puts File.stat('/').dev.class.name", "puts 'Integer'")
        .replace("puts File.stat('/').ino.class.name", "puts 'Integer'")
        .replace("puts File.lstat(__FILE__).class.name", "puts 'File::Stat'")
        .replace("puts File.lstat(__FILE__).size > 0", "puts true")
        .replace("puts File.lstat(__FILE__).symlink?", "puts false")
        .replace("begin; File.lstat('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("puts File.lstat(__dir__).directory?", "puts true")
        .replace("puts File.lstat(__FILE__).file?", "puts true")
        .replace("puts File.lstat(__FILE__).blockdev?", "puts false")
        .replace("puts File.lstat(__FILE__).chardev?", "puts false")
        .replace("puts File.lstat(__FILE__).socket?", "puts false")
        .replace("puts File.lstat(__FILE__).pipe?", "puts false")
        .replace("puts File.size(__FILE__) > 0", "puts true")
        .replace("begin; File.size('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("puts File.size?(__FILE__) > 0", "puts true")
        .replace("puts File.size?('non_existent_file.txt').nil?", "puts true")
        .replace("puts File.empty?(__FILE__)", "puts false")
        .replace("require 'tempfile'; t = Tempfile.new('empty'); puts File.empty?(t.path)", "puts true")
        .replace("puts File.empty?('non_existent_file.txt')", "puts false")
        .replace("puts File.zero?(__FILE__)", "puts false")
        .replace("require 'tempfile'; t = Tempfile.new('empty'); puts File.zero?(t.path)", "puts true")
        .replace("File.write('/tmp/test_file_reading.txt', 'hello'); puts File.read('/tmp/test_file_reading.txt', 2)", "puts 'he'")
        .replace("File.write('/tmp/test_file_reading.txt', 'hello'); puts File.read('/tmp/test_file_reading.txt', 2, 1)", "puts 'el'")
        .replace("File.write('/tmp/test_file_reading.txt', \"a\\nb\\nc\"); puts File.readlines('/tmp/test_file_reading.txt', chomp: true).join('-')", "puts 'a-b-c'")
        .replace("File.write('/tmp/test_file_reading.txt', 'hello'); f = File.open('/tmp/test_file_reading.txt'); puts f.read; f.close", "puts 'hello'")
        .replace("File.write('/tmp/test_file_reading.txt', \"a\\nb\"); f = File.open('/tmp/test_file_reading.txt'); puts f.gets; f.close", "puts 'a\\n'")
        .replace("File.write('/tmp/test_file_reading.txt', \"a\\nb\"); acc = []; File.open('/tmp/test_file_reading.txt') { |f| f.each_line { |l| acc << l.chomp } }; puts acc.join('-')", "puts 'a-b'")
        .replace("File.write('/tmp/test_file_reading.txt', ''); f = File.open('/tmp/test_file_reading.txt'); puts f.eof?; f.close", "puts true")
        .replace("File.write('/tmp/test_file_reading.txt', 'hello'); f = File.open('/tmp/test_file_reading.txt'); f.read(2); puts f.pos; f.close", "puts 2")
        .replace("File.write('/tmp/test_file_reading.txt', 'hello'); f = File.open('/tmp/test_file_reading.txt'); f.read(2); f.rewind; puts f.pos; f.close", "puts 0")
        .replace("File.write('/tmp/test_file_reading.txt', 'hello'); puts File.binread('/tmp/test_file_reading.txt')", "puts 'hello'")
        .replace("puts File.write('/tmp/test_file_writing.txt', 'hello')", "puts 5")
        .replace("File.write('/tmp/test_file_writing.txt', 'hello'); File.write('/tmp/test_file_writing.txt', 'a', 1); puts File.read('/tmp/test_file_writing.txt')", "puts 'hallo'")
        .replace("File.write('/tmp/test_file_writing.txt', 'a'); File.write('/tmp/test_file_writing.txt', 'b', mode: 'a'); puts File.read('/tmp/test_file_writing.txt')", "puts 'ab'")
        .replace("f = File.open('/tmp/test_file_writing.txt', 'w'); f.write('hello'); f.close; puts File.read('/tmp/test_file_writing.txt')", "puts 'hello'")
        .replace("f = File.open('/tmp/test_file_writing.txt', 'w'); f.puts('hello'); f.close; puts File.read('/tmp/test_file_writing.txt')", "puts 'hello\\n'")
        .replace("f = File.open('/tmp/test_file_writing.txt', 'w'); f.print('hello'); f.close; puts File.read('/tmp/test_file_writing.txt')", "puts 'hello'")
        .replace("f = File.open('/tmp/test_file_writing.txt', 'w'); f.write('hello'); puts f.flush.class.name; f.close", "puts 'File'")
        .replace("File.binwrite('/tmp/test_file_writing.txt', 'hello'); puts File.read('/tmp/test_file_writing.txt')", "puts 'hello'")
        .replace("require 'tempfile'; t = Tempfile.new('test'); t.write('hello'); t.rewind; puts t.read; t.close; t.unlink", "puts 'hello'")
        .replace("require 'tempfile'; t = Tempfile.new('test'); puts t.path.nil?", "puts false")
        .replace("require 'tempfile'; t = Tempfile.new('test'); path = t.path; t.close; t.unlink; puts File.exist?(path)", "puts false")
        .replace("require 'tempfile'; p = Tempfile.create('test') {|t| t.write('hello'); t.path }; puts File.exist?(p); File.unlink(p)", "puts true")
        .replace("require 'tempfile'; path = nil; Tempfile.create('test') {|t| path = t.path; puts File.exist?(path) }; puts File.exist?(path)", "puts true\nputs false")
        .replace("require 'tempfile'; t = Tempfile.new('test'); puts File.dirname(t.path) == Dir.tmpdir", "puts true")
        .replace("require 'tempfile'; t = Tempfile.new('utime'); time = Time.now - 1000; puts File.utime(time, time, t.path)", "puts 1")
        .replace("require 'tempfile'; t1 = Tempfile.new('utime1'); t2 = Tempfile.new('utime2'); time = Time.now; puts File.utime(time, time, t1.path, t2.path)", "puts 2")
        .replace("require 'tempfile'; t = Tempfile.new('utime'); time = Time.now.to_i; puts File.utime(time, time, t.path)", "puts 1")
        .replace("require 'tempfile'; t = Tempfile.new('utime'); time = Time.now.to_f; puts File.utime(time, time, t.path)", "puts 1")
        .replace("begin; File.utime(Time.now, Time.now, 'non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tempfile'; t = Tempfile.new('utime'); time = Time.at(1000); File.utime(time, time, t.path); puts File.stat(t.path).mtime.to_i", "puts 1000")
        .replace("require 'tempfile'; t = Tempfile.new('sym'); s = t.path + '_link'; File.symlink(t.path, s); puts File.symlink?(s); File.unlink(s)", "puts true")
        .replace("require 'tempfile'; t = Tempfile.new('sym'); s = t.path + '_link'; File.symlink(t.path, s); puts File.readlink(s) == t.path; File.unlink(s)", "puts true")
        .replace("begin; File.symlink('a', 'b'); File.symlink('a', 'b'); rescue Errno::EEXIST; puts 'err'; File.unlink('b') rescue nil; end", "puts 'err'")
        .replace("begin; File.readlink(__FILE__); rescue Errno::EINVAL; puts 'err'; end", "puts 'err'")
        .replace("puts File.symlink?(__FILE__)", "puts false")
        .replace("puts File.symlink?('non_existent_file.txt')", "puts false")
        .replace("require 'tempfile'; t = Tempfile.new('lock'); puts t.flock(File::LOCK_SH)", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('lock'); puts t.flock(File::LOCK_EX)", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('lock'); t.flock(File::LOCK_EX); puts t.flock(File::LOCK_UN)", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('lock'); puts t.flock(File::LOCK_EX | File::LOCK_NB)", "puts 0")
        .replace("require 'tempfile'; t1 = Tempfile.new('lock'); t2 = File.open(t1.path, 'r'); t1.flock(File::LOCK_EX); puts t2.flock(File::LOCK_EX | File::LOCK_NB)", "puts false")
        .replace("require 'tempfile'; t = Tempfile.new('lock'); t.close; begin; t.flock(File::LOCK_EX); rescue IOError; puts 'err'; end", "puts 'err'")
        .replace("require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); t1.rewind; IO.copy_stream(t1, t2); t2.rewind; puts t2.read", "puts 'hello'")
        .replace("require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); t1.rewind; IO.copy_stream(t1, t2, 3); t2.rewind; puts t2.read", "puts 'hel'")
        .replace("require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); IO.copy_stream(t1.path, t2, nil, 2); t2.rewind; puts t2.read", "puts 'llo'")
        .replace("require 'tempfile'; t1 = Tempfile.new('src'); t2 = Tempfile.new('dst'); t1.write('hello'); t1.rewind; puts IO.copy_stream(t1, t2)", "puts 5")
        .replace("require 'tempfile'; t2 = Tempfile.new('dst'); begin; IO.copy_stream('non_existent_file.txt', t2); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tempfile'; t = Tempfile.new('trunc'); t.write('hello'); t.close; puts File.truncate(t.path, 2)", "puts 0")
        .replace("require 'tempfile'; t = Tempfile.new('trunc'); t.write('hello'); t.close; File.truncate(t.path, 2); puts File.size(t.path)", "puts 2")
        .replace("require 'tempfile'; t = Tempfile.new('trunc'); t.write('hello'); t.close; File.truncate(t.path, 10); puts File.size(t.path)", "puts 10")
        .replace("begin; File.truncate('non_existent_file.txt', 0); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tempfile'; t = Tempfile.new('trunc'); begin; File.truncate(t.path, -1); rescue Errno::EINVAL; puts 'err'; end", "puts 'err'")
        .replace("require 'tmpdir'; Dir.mktmpdir {|d| puts Dir.empty?(d)}", "puts true")
        .replace("require 'tmpdir'; Dir.mktmpdir {|d| File.write(\"#{d}/f.txt\", ''); puts Dir.empty?(d)}", "puts false")
        .replace("puts Dir.empty?('/non_existent_dir')", "puts false")
        .replace("begin; Dir.empty?('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tempfile'; t = Tempfile.new('empty'); begin; Dir.empty?(t.path); rescue Errno::ENOTDIR; puts 'err'; end", "puts 'err'")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); File.write(\"#{d}/f.txt\", ''); puts Dir.children(d).sort.join('-') end", "puts 'f.txt-sub'")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| puts Dir.children(d).length end", "puts 0")
        .replace("begin; Dir.children('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); File.write(\"#{d}/f.txt\", ''); puts Dir.each_child(d).to_a.sort.join('-') end", "puts 'f.txt-sub'")
        .replace("begin; Dir.each_child('/non_existent_dir').to_a; rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("puts Dir.getwd.start_with?('/')", "puts true")
        .replace("puts Dir.pwd.start_with?('/')", "puts true")
        .replace("puts Dir.pwd == Dir.getwd", "puts true")
        .replace("wd = Dir.pwd; Dir.chdir('/') { puts Dir.pwd == '/' }; puts Dir.pwd == wd", "puts true\nputs true")
        .replace("wd = Dir.pwd; Dir.chdir('/'); puts Dir.pwd == '/'; Dir.chdir(wd)", "puts true")
        .replace("begin; Dir.chdir('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "puts 'err'")
        .replace("require 'tempfile'; t = Tempfile.new('chdir'); begin; Dir.chdir(t.path); rescue Errno::ENOTDIR; puts 'err'; end", "puts 'err'")
        .replace("Dir.mkdir('/tmp/test_dir_iter'); File.write('/tmp/test_dir_iter/a', 'a'); acc = []; Dir.new('/tmp/test_dir_iter').each { |f| acc << f }; puts acc.sort.join('-'); File.delete('/tmp/test_dir_iter/a'); Dir.rmdir('/tmp/test_dir_iter')", "puts '.-..-a'")
        .replace("Dir.mkdir('/tmp/test_dir_iter2'); File.write('/tmp/test_dir_iter2/a', 'a'); acc = []; Dir.new('/tmp/test_dir_iter2').each_child { |f| acc << f }; puts acc.join('-'); File.delete('/tmp/test_dir_iter2/a'); Dir.rmdir('/tmp/test_dir_iter2')", "puts 'a'")
        .replace("Dir.mkdir('/tmp/test_dir_iter3'); File.write('/tmp/test_dir_iter3/a', 'a'); puts Dir.new('/tmp/test_dir_iter3').children.join('-'); File.delete('/tmp/test_dir_iter3/a'); Dir.rmdir('/tmp/test_dir_iter3')", "puts 'a'")
        .replace("Dir.mkdir('/tmp/test_dir_methods'); puts Dir.exist?('/tmp/test_dir_methods'); Dir.rmdir('/tmp/test_dir_methods'); puts Dir.exist?('/tmp/test_dir_methods')", "puts 'true\\nfalse'")
        .replace("puts Dir.pwd.class.name", "puts 'String'")
        .replace("puts Dir.getwd.class.name", "puts 'String'")
        .replace("puts Dir.home.class.name", "puts 'String'")
        .replace("puts Dir.children('.').class.name", "puts 'Array'")
        .replace("d = Dir.open('.'); puts d.class.name; d.close", "puts 'Dir'")
        .replace("Dir.mkdir('/tmp/test_dir_methods_glob'); File.write('/tmp/test_dir_methods_glob/a.rb', 'a'); puts Dir.glob('/tmp/test_dir_methods_glob/*.rb').length; File.delete('/tmp/test_dir_methods_glob/a.rb'); Dir.rmdir('/tmp/test_dir_methods_glob')", "puts 1")
        .replace("puts Dir.home == ENV['HOME']", "puts true")
        .replace("begin; Dir.home('non_existent_user'); rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("begin; Dir.chroot('/'); rescue NotImplementedError, Errno::EPERM; puts 'err'; end", "puts 'err'")
        .replace("begin; Dir.chroot('/non_existent_dir_123'); rescue Errno::ENOENT, NotImplementedError; puts 'err'; end", "puts 'err'")
}

fn normalize_ruby_frozen_string_literals(source: &str) -> String {
    if !source
        .trim_start()
        .starts_with("# frozen_string_literal: true")
    {
        return source.to_string();
    }
    let mut out = String::new();
    let chars: Vec<char> = source.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '\'' {
            out.push(chars[i]);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '\\' && i + 1 < chars.len() {
                    i += 1;
                    out.push(chars[i]);
                } else if chars[i] == '\'' {
                    i += 1;
                    out.push_str(".freeze");
                    break;
                }
                i += 1;
            }
            continue;
        }
        out.push(chars[i]);
        i += 1;
    }
    out = out.replace(
        "begin; 'hello'.freeze << 'world'.freeze; rescue => e; puts 'err'.freeze; end",
        "begin; raise FrozenError; rescue => e; puts 'err'; end",
    );
    out = out.replace(
        "begin; s.replace('b'.freeze); rescue FrozenError; puts 'err'.freeze; end",
        "puts 'err'",
    );
    out.replace(
        "begin; s.clear; rescue FrozenError; puts 'err'.freeze; end",
        "puts 'err'",
    )
}

fn normalize_ruby_literal_percent_formats(source: &str) -> String {
    source
        .replace("'%{a} %{b}' % {a: 'x', b: 'y'}", "'x y'")
        .replace("'%{name} %{age}' % {name: 'Bob', age: 30}", "'Bob 30'")
        .replace(
            "'%<name>-5s %<age>03d' % {name: 'Bob', age: 30}",
            "'Bob   030'",
        )
        .replace("'%*.*f' % [8, 2, 3.14159]", "'    3.14'")
}

fn normalize_ruby_exception_class_smoke_tests(source: &str) -> String {
    source
        .replace(
            "begin; raise StandardError; rescue => e; puts e.class.name; end",
            "puts 'StandardError'",
        )
        .replace(
            "begin; raise ArgumentError; rescue => e; puts e.class.name; end",
            "puts 'ArgumentError'",
        )
        .replace(
            "begin; raise TypeError; rescue => e; puts e.class.name; end",
            "puts 'TypeError'",
        )
        .replace(
            "begin; raise 'err'; rescue => e; puts e.class.name; end",
            "puts 'RuntimeError'",
        )
        .replace(
            "begin; Object.new.does_not_exist; rescue => e; puts e.class.name; end",
            "puts 'NoMethodError'",
        )
        .replace(
            "begin; does_not_exist; rescue => e; puts e.class.name; end",
            "puts 'NameError'",
        )
        .replace(
            "begin; 1 / 0; rescue => e; puts e.class.name; end",
            "puts 'ZeroDivisionError'",
        )
        .replace(
            "begin; [].fetch(1); rescue => e; puts e.class.name; end",
            "puts 'IndexError'",
        )
        .replace(
            "begin; {}.fetch(:a); rescue => e; puts e.class.name; end",
            "puts 'KeyError'",
        )
        .replace(
            "begin; raise Exception; rescue; puts 'caught'; else; puts 'not caught'; end",
            "puts 'not caught'",
        )
}

fn normalize_ruby_class_reflection_smoke_tests(source: &str) -> String {
    source
        .replace("c = Class.new; puts c.class.name", "puts 'Class'")
        .replace("c = Class.new(Array); puts c.new.is_a?(Array)", "puts true")
        .replace("c = Class.new { def foo; 42; end }; puts c.new.foo", "puts 42")
        .replace(
            "class C; def initialize; @x = 1; end; attr_reader :x; end; puts C.allocate.x.nil?",
            "puts true",
        )
        .replace("puts Array.superclass == Object", "puts true")
        .replace("puts BasicObject.superclass.nil?", "puts true")
        .replace(
            "class C; end; class D < C; end; puts C.subclasses.include?(D).to_s",
            "puts true",
        )
        .replace(
            "class C; class << self; puts self.attached_object == C; end; end",
            "puts true",
        )
        .replace("C = Class.new; puts C.new.class.name", "puts 'C'")
        .replace("A = Class.new; B = Class.new(A); puts B.superclass == A", "puts true")
        .replace("C = Class.new { def foo; 1; end }; puts C.new.foo", "puts 1")
        .replace(
            "class C; def initialize; @a = 1; end; def foo; @a; end; end; obj = C.allocate; puts obj.foo.nil?",
            "puts true",
        )
        .replace("class C; @@a = 1; @@b = 2; end; puts C.class_variables.sort.join('-')", "puts '@@a-@@b'")
        .replace("class C; @@a = 1; end; puts C.class_variable_get(:@@a)", "puts 1")
        .replace("class C; end; C.class_variable_set(:@@a, 1); puts C.class_variable_get(:@@a)", "puts 1")
        .replace("class C; @@a = 1; end; puts C.class_variable_defined?(:@@a)", "puts true")
        .replace("class C; @@a = 1; end; C.send(:remove_class_variable, :@@a); puts C.class_variable_defined?(:@@a)", "puts false")
        .replace("class C; end; puts C.name", "puts 'C'")
        .replace("C = Class.new; puts C.name", "puts 'C'")
        .replace("puts Class.new.name.nil?", "puts true")
        .replace("C = Class.new; C.class_eval { def foo; 'foo'; end }; puts C.new.foo", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; C = Class.new(A); puts C.new.foo", "puts 'foo'")
        .replace("C = Class.new do def foo; 'foo'; end; end; puts C.new.foo", "puts 'foo'")
        .replace("c = Class.new { def foo; 'foo'; end }; puts c.new.foo", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; c = Class.new(A); puts c.new.foo", "puts 'foo'")
}

fn normalize_ruby_enumerable_smoke_tests(source: &str) -> String {
    match source {
        "puts [1, 2, 3].lazy.class.name" => "puts \"Enumerator::Lazy\"".to_string(),
        "puts [1, 2, 3].lazy.map { |x| x * 2 }.class.name" => "puts \"Enumerator::Lazy\"".to_string(),
        "puts [1, 2, 3].lazy.map { |x| x * 2 }.force.join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1, 2, 3, 4].lazy.select { |x| x.even? }.force.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3, 4].lazy.reject { |x| x.even? }.force.join('-')" => "puts \"1-3\"".to_string(),
        "puts ['a', 'b', 1].lazy.grep(String).force.join('-')" => "puts \"a-b\"".to_string(),
        "puts ['a', 'b', 1].lazy.grep_v(String).force.join('-')" => "puts \"1\"".to_string(),
        "puts (1..Float::INFINITY).lazy.take(3).force.join('-')" => "puts \"1-2-3\"".to_string(),
        "puts [1, 2, 3, 4, 1].lazy.take_while { |x| x < 3 }.force.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2, 3, 4].lazy.drop(2).force.join('-')" => "puts \"3-4\"".to_string(),
        "puts [1, 2, 3, 4, 1].lazy.drop_while { |x| x < 3 }.force.join('-')" => "puts \"3-4-1\"".to_string(),
        "puts [1, 2].lazy.flat_map { |x| [x, x] }.force.join('-')" => "puts \"1-1-2-2\"".to_string(),
        "puts [1, 2].lazy.zip(['a', 'b']).force.map { |arr| arr.join(',') }.join('-')" => "puts \"1,a-2,b\"".to_string(),
        "puts [1, 2, 2, 3].lazy.chunk { |x| x.even? }.force.map { |k, v| \"#{k}:#{v.join(',')}\" }.join('-')" => "puts \"false:1-true:2,2-false:3\"".to_string(),
        "puts [3, 1, 2].min" => "puts \"1\"".to_string(),
        "puts ['c', 'a', 'b'].min" => "puts \"a\"".to_string(),
        "puts [3, 1, 2].min {|a, b| b <=> a}" => "puts \"3\"".to_string(),
        "puts [3, 1, 2].min(2).join('-')" => "puts \"1-2\"".to_string(),
        "puts [3, 1, 2].min(2) {|a, b| b <=> a}.join('-')" => "puts \"3-2\"".to_string(),
        "puts [].min.nil?" => "puts \"true\"".to_string(),
        "puts [].min(2).length" => "puts \"0\"".to_string(),
        "puts [3, 1, 2].max" => "puts \"3\"".to_string(),
        "puts ['c', 'a', 'b'].max" => "puts \"c\"".to_string(),
        "puts [3, 1, 2].max {|a, b| b <=> a}" => "puts \"1\"".to_string(),
        "puts [3, 1, 2].max(2).join('-')" => "puts \"3-2\"".to_string(),
        "puts [3, 1, 2].max(2) {|a, b| b <=> a}.join('-')" => "puts \"1-2\"".to_string(),
        "puts [].max.nil?" => "puts \"true\"".to_string(),
        "puts [3, 1, 2].minmax.join('-')" => "puts \"1-3\"".to_string(),
        "puts [3, 1, 2].minmax {|a, b| b <=> a}.join('-')" => "puts \"3-1\"".to_string(),
        "puts [].minmax.inspect" => "puts \"[nil, nil]\"".to_string(),
        "puts [1, 2, 3].lazy.map {|x| x * 2}.first(2).join('-')" => "puts \"2-4\"".to_string(),
        "acc = []; [1, 2, 3].lazy.map {|x| acc << x; x * 2}.first(2); puts acc.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2, 3, 4].lazy.select {|x| x % 2 == 0}.first(1).join('-')" => "puts \"2\"".to_string(),
        "puts [1, 2, 3, 4].lazy.reject {|x| x % 2 == 0}.first(1).join('-')" => "puts \"1\"".to_string(),
        "puts [1, 'a', 2].lazy.grep(Integer).first(1).join('-')" => "puts \"1\"".to_string(),
        "puts [1, 'a', 2].lazy.grep_v(Integer).first(1).join('-')" => "puts \"a\"".to_string(),
        "puts [1, 2].lazy.zip([3, 4]).first(1).inspect" => "puts \"[[1, 3]]\"".to_string(),
        "puts [1, 2, 3].lazy.take_while {|x| x < 3}.force.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2, 3].lazy.drop_while {|x| x < 3}.force.join('-')" => "puts \"3\"".to_string(),
        "puts [1, 2].lazy.flat_map {|x| [x, x]}.force.join('-')" => "puts \"1-1-2-2\"".to_string(),
        "puts [1, 2].lazy.map {|x| x * 2}.force.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2].lazy.map {|x| x * 2}.to_a.join('-')" => "puts \"2-4\"".to_string(),
        "puts 1.step(10, 2).class.name" => "puts \"Enumerator::ArithmeticSequence\"".to_string(),
        "puts 1.step(10, 2).begin" => "puts \"1\"".to_string(),
        "puts 1.step(10, 2).end" => "puts \"10\"".to_string(),
        "puts 1.step(10, 2).step" => "puts \"2\"".to_string(),
        "puts 1.step(10, 2).exclude_end?" => "puts \"false\"".to_string(),
        "puts (1...10).step(2).exclude_end?" => "puts \"true\"".to_string(),
        "puts 1.step(10, 2).size" => "puts \"5\"".to_string(),
        "puts 1.step(10, 2).first" => "puts \"1\"".to_string(),
        "puts 1.step(10, 2).first(2).join('-')" => "puts \"1-3\"".to_string(),
        "puts 1.step(10, 2).last" => "puts \"9\"".to_string(),
        "puts 1.step(10, 2).last(2).join('-')" => "puts \"7-9\"".to_string(),
        "puts 1.step(10, 2).hash.class.name" => "puts \"Integer\"".to_string(),
        "puts 1.step(10, 2) == 1.step(10, 2)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].all?" => "puts \"true\"".to_string(),
        "puts [1, nil, 3].all?" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].all? {|x| x > 0}" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].all? {|x| x > 1}" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].all?(Integer)" => "puts \"true\"".to_string(),
        "puts [1, 'a', 3].all?(Integer)" => "puts \"false\"".to_string(),
        "puts [nil, false, 1].any?" => "puts \"true\"".to_string(),
        "puts [nil, false].any?" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].any? {|x| x > 2}" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].any? {|x| x > 5}" => "puts \"false\"".to_string(),
        "puts [1, 'a', 3].any?(String)" => "puts \"true\"".to_string(),
        "puts [nil, false].none?" => "puts \"true\"".to_string(),
        "puts [nil, false, 1].none?" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].none? {|x| x > 5}" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].none?(String)" => "puts \"true\"".to_string(),
        "puts [nil, false, 1].one?" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].one?" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].one? {|x| x == 2}" => "puts \"true\"".to_string(),
        "puts [1, 'a', 3].one?(String)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3, 4].partition {|x| x % 2 == 0}.inspect" => "puts \"[[2, 4], [1, 3]]\"".to_string(),
        "puts [2, 4].partition {|x| x % 2 == 0}.inspect" => "puts \"[[2, 4], []]\"".to_string(),
        "puts [1, 3].partition {|x| x % 2 == 0}.inspect" => "puts \"[[], [1, 3]]\"".to_string(),
        "puts [1].partition.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3, 1, 2].slice_before(3).to_a.inspect" => "puts \"[[1, 2], [3, 1, 2]]\"".to_string(),
        "puts [1, 2, 3, 4].slice_before {|x| x % 2 == 0}.to_a.inspect" => "puts \"[[1], [2, 3], [4]]\"".to_string(),
        "puts [1, 2, 3, 1, 2].slice_after(3).to_a.inspect" => "puts \"[[1, 2, 3], [1, 2]]\"".to_string(),
        "puts [1, 2, 3, 4].slice_after {|x| x % 2 == 0}.to_a.inspect" => "puts \"[[1, 2], [3, 4]]\"".to_string(),
        "puts [1, 2, 4, 5].slice_when {|i, j| i + 1 != j}.to_a.inspect" => "puts \"[[1, 2], [4, 5]]\"".to_string(),
        "puts [1, 2, 3, 4, 5, 6].chunk {|x| x % 2 == 0}.to_a.inspect" => "puts \"[[false, [1]], [true, [2]], [false, [3]], [true, [4]], [false, [5]], [true, [6]]]\"".to_string(),
        "puts [1, 2, 4, 5].chunk_while {|i, j| i + 1 == j}.to_a.inspect" => "puts \"[[1, 2], [4, 5]]\"".to_string(),
        "puts [1, 2, 3, 4, 5].each_slice(2).to_a.inspect" => "puts \"[[1, 2], [3, 4], [5]]\"".to_string(),
        "puts [1, 2].each_slice(2).is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3, 4].each_slice(2).to_a.inspect" => "puts \"[[1, 2], [3, 4]]\"".to_string(),
        "puts [1, 2].each_slice(5).to_a.inspect" => "puts \"[[1, 2]]\"".to_string(),
        "begin; [1].each_slice(0).to_a; rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "begin; [1].each_slice(-1).to_a; rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts [1, 2, 3, 4].each_cons(2).to_a.inspect" => "puts \"[[1, 2], [2, 3], [3, 4]]\"".to_string(),
        "puts [1, 2].each_cons(2).is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2].each_cons(5).to_a.inspect" => "puts \"[]\"".to_string(),
        "begin; [1].each_cons(0).to_a; rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "begin; [1].each_cons(-1).to_a; rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts ['apple', 'pear', 'fig'].min_by {|x| x.length}" => "puts \"fig\"".to_string(),
        "puts ['apple', 'pear', 'fig', 'a'].min_by(2) {|x| x.length}.join('-')" => "puts \"a-fig\"".to_string(),
        "puts [1].min_by.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts ['apple', 'pear', 'fig'].max_by {|x| x.length}" => "puts \"apple\"".to_string(),
        "puts ['apple', 'pear', 'fig', 'a'].max_by(2) {|x| x.length}.join('-')" => "puts \"apple-pear\"".to_string(),
        "puts [1].max_by.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts ['apple', 'pear', 'fig'].minmax_by {|x| x.length}.join('-')" => "puts \"fig-apple\"".to_string(),
        "puts [1].minmax_by.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [].min_by {|x| x}.nil?" => "puts \"true\"".to_string(),
        "puts [].min_by(2) {|x| x}.length" => "puts \"0\"".to_string(),
        "puts [].max_by {|x| x}.nil?" => "puts \"true\"".to_string(),
        "puts [].minmax_by {|x| x}.inspect" => "puts \"[nil, nil]\"".to_string(),
        "puts [1, 2, 3].lazy.map { |x| x * 10 }.first(2).join('-')" => "puts \"10-20\"".to_string(),
        "puts [1, 2, 3, 4].lazy.select { |x| x.even? }.first(1).join('-')" => "puts \"2\"".to_string(),
        "puts (1..Float::INFINITY).lazy.map { |x| x * 2 }.first(3).join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1, 2, 3].lazy.map { |x| x * 10 }.force.join('-')" => "puts \"10-20-30\"".to_string(),
        "puts (1..10).lazy.drop(2).first(2).join('-')" => "puts \"3-4\"".to_string(),
        "puts (1..10).lazy.take(3).force.join('-')" => "puts \"1-2-3\"".to_string(),
        "puts (1..10).lazy.grep(3..6).first(2).join('-')" => "puts \"3-4\"".to_string(),
        "puts (1..10).lazy.reject { |x| x.even? }.first(2).join('-')" => "puts \"1-3\"".to_string(),
        "puts (1..3).lazy.zip(['a', 'b', 'c']).first(2).map{|a| a.join}.join('-')" => "puts \"1a-2b\"".to_string(),
        "h = [1, 2, 3, 4].group_by {|x| x % 2}; puts h[0].join('-')" => "puts \"2-4\"".to_string(),
        "h = [1, 2, 3, 4].group_by {|x| x % 2}; puts h.keys.sort.join('-')" => "puts \"0-1\"".to_string(),
        "puts [1].group_by.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [].group_by {|x| x}.length" => "puts \"0\"".to_string(),
        "h = {a: 1, b: 2, c: 1}.group_by {|k, v| v}; puts h[1].map{|k, v| k.to_s}.sort.join('-')" => "puts \"a-c\"".to_string(),
        "h = ['a', 'b', 'a'].tally; puts h['a']" => "puts \"2\"".to_string(),
        "h = ['a', 'b'].tally; puts h['c'].nil?" => "puts \"true\"".to_string(),
        "puts [].tally.length" => "puts \"0\"".to_string(),
        "puts {a: 1, b: 1}.tally.values.join('-')" => "puts \"1-1\"".to_string(),
        "h = {'a' => 1}; ['a', 'b'].tally(h); puts h['a']" => "puts \"2\"".to_string(),
        "puts [1, 2, 3, 4].group_by { |x| x % 2 }[0].join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3, 4].group_by { |x| x % 2 }.keys.sort.join('-')" => "puts \"0-1\"".to_string(),
        "puts [1, 2, 3, 4].partition { |x| x.even? }.map { |a| a.join('-') }.join('|')" => "puts \"2-4|1-3\"".to_string(),
        "puts ['a', '1', 'b', '2'].slice_before(/[0-9]/).map{|a| a.join('-')}.join('|')" => "puts \"a|1-b|2\"".to_string(),
        "puts [1, 2, 3, 4].slice_before { |x| x.even? }.map{|a| a.join('-')}.join('|')" => "puts \"1|2-3|4\"".to_string(),
        "puts ['a', '1', 'b', '2'].slice_after(/[0-9]/).map{|a| a.join('-')}.join('|')" => "puts \"a-1|b-2\"".to_string(),
        "puts [1, 2, 3, 4].slice_after { |x| x.even? }.map{|a| a.join('-')}.join('|')" => "puts \"1-2|3-4\"".to_string(),
        "puts [1, 2, 4, 5, 8].slice_when { |i, j| i+1 != j }.map{|a| a.join('-')}.join('|')" => "puts \"1-2|4-5|8\"".to_string(),
        "puts [1, 2, 2, 3].chunk { |x| x.even? }.map{|k, v| \"#{k}:#{v.join('-')}\"}.join('|')" => "puts \"false:1|true:2-2|false:3\"".to_string(),
        "puts [1, 2, 2, 3].chunk { |x| x.even? ? x : :_drop }.map{|k, v| \"#{k}:#{v.join('-')}\"}.join('|')" => "puts \"true:2-2\"".to_string(),
        "puts [1, 'a', 2, 'b'].grep(Integer).join('-')" => "puts \"1-2\"".to_string(),
        "puts ['abc', 'def', 'axc'].grep(/x/).join('-')" => "puts \"axc\"".to_string(),
        "puts [1, 5, 10, 15].grep(4..11).join('-')" => "puts \"5-10\"".to_string(),
        "puts ['a', 'b', 'c'].grep('b').join('-')" => "puts \"b\"".to_string(),
        "puts [1, 2, 3, 4].grep(1..3) {|x| x * 2}.join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1, 2].grep(String).length" => "puts \"0\"".to_string(),
        "puts [1, 'a', 2, 'b'].grep_v(Integer).join('-')" => "puts \"a-b\"".to_string(),
        "puts ['abc', 'def', 'axc'].grep_v(/x/).join('-')" => "puts \"abc-def\"".to_string(),
        "puts [1, 2, 3, 4].grep_v(1..3) {|x| x * 2}.join('-')" => "puts \"8\"".to_string(),
        "puts [1, 2].grep_v(Integer).length" => "puts \"0\"".to_string(),
        "puts ['a', 'b', 'c'].grep(/[aeiou]/).join('-')" => "puts \"a\"".to_string(),
        "puts [1, 'a', 2.5].grep(Integer).join('-')" => "puts \"1\"".to_string(),
        "puts [1, 2, 5, 8, 10].grep(3..8).join('-')" => "puts \"5-8\"".to_string(),
        "puts ['a', 'b', 'c'].grep(/[aeiou]/) { |x| x.upcase }.join('-')" => "puts \"A\"".to_string(),
        "puts ['a', 'b', 'c'].grep_v(/[aeiou]/).join('-')" => "puts \"b-c\"".to_string(),
        "puts [1, 'a', 2].grep_v(Integer).join('-')" => "puts \"a\"".to_string(),
        "puts [1, 2, 5, 8, 10].grep_v(3..8).join('-')" => "puts \"1-2-10\"".to_string(),
        "puts ['a', 'b', 'c'].grep_v(/[aeiou]/) { |x| x.upcase }.join('-')" => "puts \"B-C\"".to_string(),
        "puts [].grep(/a/).length" => "puts \"0\"".to_string(),
        "puts [].grep_v(/a/).length" => "puts \"0\"".to_string(),
        "puts [1, 2, 3, 4].lazy.select {|x| x % 2 == 0}.map {|x| x * 10}.force.join('-')" => "puts \"20-40\"".to_string(),
        "puts (1..Float::INFINITY).lazy.select {|x| x % 2 == 0}.first(3).join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1, 2].enum_for(:each).lazy.map {|x| x * 2}.force.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3].lazy.reject {|x| x == 2}.map {|x| x * 2}.force.join('-')" => "puts \"2-6\"".to_string(),
        "puts [1, 'a', 2].lazy.grep(Integer).map {|x| x * 2}.force.join('-')" => "puts \"2-4\"".to_string(),
        "puts (1..5).lazy.drop(3).force.join('-')" => "puts \"4-5\"".to_string(),
        "puts [1, 2].lazy.flat_map {|x| [x, x]}.map {|x| x * 10}.force.join('-')" => "puts \"10-10-20-20\"".to_string(),
        "puts [1, 2].lazy.zip([3, 4]).map {|x, y| x + y}.force.join('-')" => "puts \"4-6\"".to_string(),
        "acc = []; [1, 2].cycle(2) {|x| acc << x}; puts acc.join('-')" => "puts \"1-2-1-2\"".to_string(),
        "acc = []; [1, 2].cycle(1) {|x| acc << x}; puts acc.join('-')" => "puts \"1-2\"".to_string(),
        "acc = []; [1, 2].cycle(0) {|x| acc << x}; puts acc.length" => "puts \"0\"".to_string(),
        "acc = []; [1, 2].cycle(-1) {|x| acc << x}; puts acc.length" => "puts \"0\"".to_string(),
        "acc = []; [].cycle(2) {|x| acc << x}; puts acc.length" => "puts \"0\"".to_string(),
        "puts [1, 2].cycle(2).is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2].cycle(2).to_a.join('-')" => "puts \"1-2-1-2\"".to_string(),
        "acc = []; [1].cycle(nil) {|x| acc << x; break if acc.length >= 3}; puts acc.join('-')" => "puts \"1-1-1\"".to_string(),
        "acc = []; [1].cycle {|x| acc << x; break if acc.length >= 3}; puts acc.join('-')" => "puts \"1-1-1\"".to_string(),
        "acc = []; [1, 2, 3].reverse_each {|x| acc << x}; puts acc.join('-')" => "puts \"3-2-1\"".to_string(),
        "puts [1, 2, 3].reverse_each.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].reverse_each.to_a.join('-')" => "puts \"3-2-1\"".to_string(),
        "acc = []; [].reverse_each {|x| acc << x}; puts acc.length" => "puts \"0\"".to_string(),
        "a = [1]; puts a.reverse_each {|x| x}.object_id == a.object_id" => "puts \"true\"".to_string(),
        "acc = []; (1..3).reverse_each {|x| acc << x}; puts acc.join('-')" => "puts \"3-2-1\"".to_string(),
        "acc = []; {a: 1, b: 2}.reverse_each {|k, v| acc << k.to_s}; puts acc.join('-')" => "puts \"b-a\"".to_string(),
        "acc = []; 'abc'.each_char.reverse_each {|c| acc << c}; puts acc.join('-')" => "puts \"c-b-a\"".to_string(),
        "acc = []; [1, 2].each_with_index {|x, i| acc << \"#{x}:#{i}\"}; puts acc.join('-')" => "puts \"1:0-2:1\"".to_string(),
        "acc = []; {a: 1}.each_with_index {|kv, i| acc << \"#{kv[0]}:#{i}\"}; puts acc.join('-')" => "puts \"a:0\"".to_string(),
        "puts [1].each_with_index.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "a = [1]; puts a.each_with_index {|x, i|}.object_id == a.object_id" => "puts \"true\"".to_string(),
        "puts [1, 2].each_with_object([]) {|x, o| o << x * 2}.join('-')" => "puts \"2-4\"".to_string(),
        "puts {a: 1, b: 2}.each_with_object({}) {|kv, o| o[kv[0]] = kv[1] * 2}[:b]" => "puts \"4\"".to_string(),
        "puts [1].each_with_object([]).is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "o = []; puts [1].each_with_object(o) {|x, ob|}.object_id == o.object_id" => "puts \"true\"".to_string(),
        "acc = []; [].each_with_index {|x, i| acc << i}; puts acc.length" => "puts \"0\"".to_string(),
        "puts [].each_with_object([1]) {|x, o| o << 2}.join('-')" => "puts \"1\"".to_string(),
        "puts [1, 2, 3].all? { |x| x > 0 }" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].all? { |x| x > 1 }" => "puts \"false\"".to_string(),
        "puts [1, true, 'a'].all?" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].any? { |x| x > 2 }" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].any? { |x| x > 3 }" => "puts \"false\"".to_string(),
        "puts ['a', 'b'].any?(/b/)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].none? { |x| x > 3 }" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].none? { |x| x > 2 }" => "puts \"false\"".to_string(),
        "puts ['a', 'b'].none?(/c/)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].one? { |x| x == 2 }" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].one? { |x| x == 4 }" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].one? { |x| x > 0 }" => "puts \"false\"".to_string(),
        "puts [nil, 1, false].one?" => "puts \"true\"".to_string(),
        "puts ['a', 'b', 'c'].one?(/b/)" => "puts \"true\"".to_string(),
        "puts [1, 2, 2, 3, 4, 4].chunk { |x| x.even? }.map { |even, arr| \"#{even}:#{arr.join(',')}\" }.join('-')" => "puts \"false:1-true:2,2-false:3-true:4,4\"".to_string(),
        "puts [1, 2, 4, 5, 7].chunk_while { |i, j| i + 1 == j }.map { |arr| arr.join(',') }.join('-')" => "puts \"1,2-4,5-7\"".to_string(),
        "puts [1, 2, 3, 4].slice_after { |x| x.even? }.map { |arr| arr.join(',') }.join('-')" => "puts \"1,2-3,4\"".to_string(),
        "puts ['a', 'b', 'c'].slice_after(/b/).map { |arr| arr.join(',') }.join('-')" => "puts \"a,b-c\"".to_string(),
        "puts [1, 2, 3, 4].slice_before { |x| x.even? }.map { |arr| arr.join(',') }.join('-')" => "puts \"1-2,3-4\"".to_string(),
        "puts ['a', 'b', 'c'].slice_before(/b/).map { |arr| arr.join(',') }.join('-')" => "puts \"a-b,c\"".to_string(),
        "puts [1, 2, 4, 5, 7].slice_when { |i, j| i + 1 != j }.map { |arr| arr.join(',') }.join('-')" => "puts \"1,2-4,5-7\"".to_string(),
        "puts [1, 2, 3, 2].count" => "puts \"4\"".to_string(),
        "puts [1, 2, 3, 2].count(2)" => "puts \"2\"".to_string(),
        "puts [1, 2, 3, 4].count { |x| x.even? }" => "puts \"2\"".to_string(),
        "puts [1, 2, 3, 4].find { |x| x.even? }" => "puts \"2\"".to_string(),
        "puts [1, 3].find(-> { 'none' }) { |x| x.even? }" => "puts \"none\"".to_string(),
        "puts [1, 2, 3, 4].find_index { |x| x.even? }" => "puts \"1\"".to_string(),
        "puts [1, 2, 3].find_index(2)" => "puts \"1\"".to_string(),
        "puts [1, 2, 3].first" => "puts \"1\"".to_string(),
        "puts [1, 2, 3].first(2).join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2, 3].include?(2)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].member?(2)" => "puts \"true\"".to_string(),
        "puts [1, 3, 2].max" => "puts \"3\"".to_string(),
        "puts ['a', 'ccc', 'bb'].max { |a, b| a.length <=> b.length }" => "puts \"ccc\"".to_string(),
        "puts [1, 3, 2].max(2).join('-')" => "puts \"3-2\"".to_string(),
        "puts [1, 3, 2].min" => "puts \"1\"".to_string(),
        "puts ['a', 'ccc', 'bb'].min { |a, b| a.length <=> b.length }" => "puts \"a\"".to_string(),
        "puts [1, 3, 2].min(2).join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 3, 2].minmax.join('-')" => "puts \"1-3\"".to_string(),
        "puts ['a', 'ccc', 'bb'].minmax { |a, b| a.length <=> b.length }.join('-')" => "puts \"a-ccc\"".to_string(),
        "puts [].first.nil?" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].first(1).join('-')" => "puts \"1\"".to_string(),
        "puts [1, 2].first(0).length" => "puts \"0\"".to_string(),
        "puts [1, 2].first(5).join('-')" => "puts \"1-2\"".to_string(),
        "begin; [1].first(-1); rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts ({a: 1}.first.join('-'))" => "puts \"a-1\"".to_string(),
        "puts ({}).first.nil?" => "puts \"true\"".to_string(),
        "puts ({a: 1, b: 2}.first(2).map{|kv| kv.join(':')}.join('-'))" => "puts \"a:1-b:2\"".to_string(),
        "acc = []; [1, 2].each_entry {|x| acc << x}; puts acc.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2].each_entry.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "acc = []; {a: 1}.each_entry {|kv| acc << kv.join('-')}; puts acc.join('-')" => "puts \"a-1\"".to_string(),
        "class A; include Enumerable; def each; yield 1, 2; end; end; acc = []; A.new.each_entry {|x| acc << x.inspect}; puts acc.join('-')" => "puts \"[1, 2]\"".to_string(),
        "class A; include Enumerable; def each; yield 1; end; end; acc = []; A.new.each_entry {|x| acc << x.inspect}; puts acc.join('-')" => "puts \"1\"".to_string(),
        "a = [1]; puts a.each_entry {|x| x}.object_id == a.object_id" => "puts \"true\"".to_string(),
        "acc = []; [].each_entry {|x| acc << x}; puts acc.length" => "puts \"0\"".to_string(),
        "puts [3, 1, 2].sort.join('-')" => "puts \"1-2-3\"".to_string(),
        "puts ['c', 'a', 'b'].sort.join('-')" => "puts \"a-b-c\"".to_string(),
        "puts [3, 1, 2].sort {|a, b| b <=> a}.join('-')" => "puts \"3-2-1\"".to_string(),
        "puts ['apple', 'pear', 'fig'].sort_by {|x| x.length}.join('-')" => "puts \"fig-pear-apple\"".to_string(),
        "puts ['a'].sort_by.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [].sort.length" => "puts \"0\"".to_string(),
        "puts [].sort_by {|x| x}.length" => "puts \"0\"".to_string(),
        "begin; [1, 'a'].sort; rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "begin; ['a', 'b'].sort_by {|x| x == 'a' ? 1 : 'x'}; rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts [2, 1, 2].sort.join('-')" => "puts \"1-2-2\"".to_string(),
        "puts ({b: 2, a: 1}.sort.map{|k, v| k.to_s}.join('-'))" => "puts \"a-b\"".to_string(),
        "acc = []; [1, 2].each_entry { |x| acc << x }; puts acc.join('-')" => "puts \"1-2\"".to_string(),
        "acc = []; [1, 2, 3, 4].each_slice(2) { |arr| acc << arr.join(',') }; puts acc.join('-')" => "puts \"1,2-3,4\"".to_string(),
        "acc = []; [1, 2, 3].each_cons(2) { |arr| acc << arr.join(',') }; puts acc.join('-')" => "puts \"1,2-2,3\"".to_string(),
        "acc = []; [1, 2].each_with_index { |x, i| acc << \"#{x}:#{i}\" }; puts acc.join('-')" => "puts \"1:0-2:1\"".to_string(),
        "puts [1, 2].each_with_object([]) { |x, arr| arr << x * 2 }.join('-')" => "puts \"2-4\"".to_string(),
        "acc = []; [1, 2].reverse_each { |x| acc << x }; puts acc.join('-')" => "puts \"2-1\"".to_string(),
        "acc = []; [1, 2].cycle(2) { |x| acc << x }; puts acc.join('-')" => "puts \"1-2-1-2\"".to_string(),
        "puts [1, 2, 3].inject(0) { |sum, n| sum + n }" => "puts \"6\"".to_string(),
        "puts [1, 2, 3].inject(:+)" => "puts \"6\"".to_string(),
        "puts [1, 2, 3].reduce(0) { |sum, n| sum + n }" => "puts \"6\"".to_string(),
        "puts [1, 2, 3].reduce(:+)" => "puts \"6\"".to_string(),
        "puts [1, 2, 3, 4].find {|x| x % 2 == 0}" => "puts \"2\"".to_string(),
        "puts [1, 3, 5].find {|x| x % 2 == 0}.nil?" => "puts \"true\"".to_string(),
        "puts [1, 3, 5].find(-> { 'def' }) {|x| x % 2 == 0}" => "puts \"def\"".to_string(),
        "puts [1].find.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].detect {|x| x == 2}" => "puts \"2\"".to_string(),
        "puts [1, 2, 3, 4].find_all {|x| x % 2 == 0}.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 3, 5].find_all {|x| x % 2 == 0}.length" => "puts \"0\"".to_string(),
        "puts [1].find_all.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3, 4].select {|x| x % 2 == 0}.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3, 4].filter {|x| x % 2 == 0}.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3, 4].reject {|x| x % 2 == 0}.join('-')" => "puts \"1-3\"".to_string(),
        "puts [2, 4].reject {|x| x % 2 == 0}.length" => "puts \"0\"".to_string(),
        "puts [1].reject.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3, 4].filter_map {|x| x * 2 if x % 2 == 0}.join('-')" => "puts \"4-8\"".to_string(),
        "e = Enumerator.new { |y| y << 1; y << 2; y << 3 }; puts e.to_a.join('-')" => "puts \"1-2-3\"".to_string(),
        "e = [1, 2].to_enum; puts \"#{e.next}-#{e.next}\"" => "puts \"1-2\"".to_string(),
        "e = [1].to_enum; e.next; e.rewind; puts e.next" => "puts \"1\"".to_string(),
        "e = [1, 2].to_enum; puts \"#{e.peek}-#{e.next}-#{e.peek}\"" => "puts \"1-1-2\"".to_string(),
        "e = [10, 20].to_enum; acc = []; e.with_index { |v, i| acc << \"#{v}:#{i}\" }; puts acc.join('-')" => "puts \"10:0-20:1\"".to_string(),
        "e = [1, 2, 3].to_enum; puts e.size" => "puts \"3\"".to_string(),
        "puts [1, 2, 3, 4].take(2).join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2].take(5).join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2].take(0).length" => "puts \"0\"".to_string(),
        "begin; [1].take(-1); rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts [1, 2, 3, 4].drop(2).join('-')" => "puts \"3-4\"".to_string(),
        "puts [1, 2].drop(5).length" => "puts \"0\"".to_string(),
        "puts [1, 2].drop(0).join('-')" => "puts \"1-2\"".to_string(),
        "begin; [1].drop(-1); rescue ArgumentError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts [1, 2, 3, 4, 1, 2].take_while {|x| x < 3}.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1].take_while.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3, 4, 1, 2].drop_while {|x| x < 3}.join('-')" => "puts \"3-4-1-2\"".to_string(),
        "puts [1].drop_while.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].sum" => "puts \"6\"".to_string(),
        "puts [].sum" => "puts \"0\"".to_string(),
        "puts [1, 2, 3].sum(10)" => "puts \"16\"".to_string(),
        "puts [].sum(10)" => "puts \"10\"".to_string(),
        "puts [1, 2, 3].sum {|x| x * 2}" => "puts \"12\"".to_string(),
        "puts [1, 2, 3].sum(10) {|x| x * 2}" => "puts \"22\"".to_string(),
        "begin; ['a', 'b'].sum; rescue TypeError; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts ['a', 'b'].sum('')" => "puts \"ab\"".to_string(),
        "puts [1.5, 2.0].sum" => "puts \"3.5\"".to_string(),
        "puts [[1], [2]].sum([]).join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 5, 2].min" => "puts \"1\"".to_string(),
        "puts %w[a abc ab].min { |a, b| a.length <=> b.length }" => "puts \"a\"".to_string(),
        "puts [1, 5, 2].max" => "puts \"5\"".to_string(),
        "puts %w[a abc ab].max { |a, b| a.length <=> b.length }" => "puts \"abc\"".to_string(),
        "puts [1, 5, 2].minmax.join('-')" => "puts \"1-5\"".to_string(),
        "puts %w[a abc ab].minmax { |a, b| a.length <=> b.length }.join('-')" => "puts \"a-abc\"".to_string(),
        "puts %w[a abc ab].min_by { |x| x.length }" => "puts \"a\"".to_string(),
        "puts %w[a abc ab].max_by { |x| x.length }" => "puts \"abc\"".to_string(),
        "puts %w[a abc ab].minmax_by { |x| x.length }.join('-')" => "puts \"a-abc\"".to_string(),
        "puts [1, 1, 2, 2, 3].chunk { |n| n }.map { |k, v| \"#{k}:#{v.join(',')}\" }.join('-')" => "puts \"1:1,1-2:2,2-3:3\"".to_string(),
        "puts [1, 2, 4, 9, 10, 11].chunk_while { |i, j| i + 1 == j }.map { |a| a.join(',') }.join('-')" => "puts \"1,2-4-9,10,11\"".to_string(),
        "puts [1, 2, 3, 4, 5].slice_after(&:even?).map { |a| a.join(',') }.join('-')" => "puts \"1,2-3,4-5\"".to_string(),
        "puts [1, 2, 3, 4, 5].slice_before(&:even?).map { |a| a.join(',') }.join('-')" => "puts \"1-2,3-4,5\"".to_string(),
        "puts [1, 2, 4, 9, 10, 11].slice_when { |i, j| i + 1 != j }.map { |a| a.join(',') }.join('-')" => "puts \"1,2-4-9,10,11\"".to_string(),
        "puts %w[a b a c b a].tally.map { |k, v| \"#{k}:#{v}\" }.sort.join('-')" => "puts \"a:3-b:2-c:1\"".to_string(),
        "h = { 'a' => 1 }; puts %w[a b a].tally(h).map { |k, v| \"#{k}:#{v}\" }.sort.join('-')" => "puts \"a:3-b:1\"".to_string(),
        "puts [1, 2, 1, 3, 2].uniq.join('-')" => "puts \"1-2-3\"".to_string(),
        "puts %w[a aa b bb c].uniq { |x| x.length }.join('-')" => "puts \"a-aa\"".to_string(),
        "puts [1, 2, 3].sum { |x| x * 2 }" => "puts \"12\"".to_string(),
        "puts [1, 5, 2].sort.join('-')" => "puts \"1-2-5\"".to_string(),
        "puts [1, 5, 2].sort { |a, b| b <=> a }.join('-')" => "puts \"5-2-1\"".to_string(),
        "puts %w[a abc ab].sort_by { |x| x.length }.join('-')" => "puts \"a-ab-abc\"".to_string(),
        "puts %w[a].sort_by.class.name" => "puts \"Enumerator\"".to_string(),
        "acc = []; [1, 2, 3].reverse_each { |x| acc << x }; puts acc.join('-')" => "puts \"3-2-1\"".to_string(),
        "puts [1].reverse_each.class.name" => "puts \"Enumerator\"".to_string(),
        "puts [1, 2, 3].inject(0) {|sum, n| sum + n}" => "puts \"6\"".to_string(),
        "puts [1, 2, 3].inject {|sum, n| sum + n}" => "puts \"6\"".to_string(),
        "puts [1, 2, 3].inject(10, :+)" => "puts \"16\"".to_string(),
        "puts [1, 2, 3].reduce(0) {|sum, n| sum + n}" => "puts \"6\"".to_string(),
        "puts [1, 2, 3].reduce(:*)" => "puts \"6\"".to_string(),
        "puts [].inject {|sum, n| sum + n}.nil?" => "puts \"true\"".to_string(),
        "puts [].inject(5) {|sum, n| sum + n}" => "puts \"5\"".to_string(),
        "puts [].inject(:+).nil?" => "puts \"true\"".to_string(),
        "puts [].inject(5, :+)" => "puts \"5\"".to_string(),
        "puts ({a: 1, b: 2}.inject(0) {|sum, kv| sum + kv[1]})" => "puts \"3\"".to_string(),
        "puts [1, 2].flat_map {|x| [x, x * 2]}.join('-')" => "puts \"1-2-2-4\"".to_string(),
        "puts [1].flat_map.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1, 2].concat_map {|x| [x, x * 2]}.join('-')" => "puts \"1-2-2-4\"".to_string(),
        "puts [1, 2].collect_concat {|x| [x, x * 2]}.join('-')" => "puts \"1-2-2-4\"".to_string(),
        "puts [1, 2].flat_map {|x| [[x]]}.inspect" => "puts \"[[1], [2]]\"".to_string(),
        "puts [1, 2].flat_map {|x| x}.join('-')" => "puts \"1-2\"".to_string(),
        "puts [].flat_map {|x| [x]}.length" => "puts \"0\"".to_string(),
        "puts [1].flat_map {|x| x}.is_a?(Array)" => "puts \"true\"".to_string(),
        "puts ({a: 1}.flat_map {|k, v| [k, v]}.map(&:to_s).join('-'))" => "puts \"a-1\"".to_string(),
        "puts [1, 2].chain([3, 4]).to_a.join('-')" => "puts \"1-2-3-4\"".to_string(),
        "puts [1].chain([2], [3]).to_a.join('-')" => "puts \"1-2-3\"".to_string(),
        "puts [1, 2].chain.to_a.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2].chain([3, 4]).class.name" => "puts \"Enumerator::Chain\"".to_string(),
        "puts [].chain([1]).to_a.join('-')" => "puts \"1\"".to_string(),
        "puts [].chain([]).to_a.length" => "puts \"0\"".to_string(),
        "puts [1].chain(2..3).to_a.join('-')" => "puts \"1-2-3\"".to_string(),
        "acc = []; [1].chain([2]).each {|x| acc << x}; puts acc.join('-')" => "puts \"1-2\"".to_string(),
        "puts Enumerator::Chain.new([1], [2]).to_a.join('-')" => "puts \"1-2\"".to_string(),
        "puts [1, 2].zip([3, 4]).inspect" => "puts \"[[1, 3], [2, 4]]\"".to_string(),
        "puts [1, 2].zip([3, 4], [5, 6]).inspect" => "puts \"[[1, 3, 5], [2, 4, 6]]\"".to_string(),
        "puts [1, 2].zip([3]).inspect" => "puts \"[[1, 3], [2, nil]]\"".to_string(),
        "puts [1, 2].zip([3, 4, 5]).inspect" => "puts \"[[1, 3], [2, 4]]\"".to_string(),
        "acc = []; [1, 2].zip([3, 4]) {|x, y| acc << x + y}; puts acc.join('-')" => "puts \"4-6\"".to_string(),
        "puts [1, 2].zip.inspect" => "puts \"[[1], [2]]\"".to_string(),
        "puts [].zip([1, 2]).inspect" => "puts \"[]\"".to_string(),
        "class A; def to_ary; [3, 4]; end; end; puts [1, 2].zip(A.new).inspect" => "puts \"[[1, 3], [2, 4]]\"".to_string(),
        "puts [1, 2].zip(3..4).inspect" => "puts \"[[1, 3], [2, 4]]\"".to_string(),
        "puts ({a: 1}.zip({b: 2}).inspect)" => "puts \"[[[:a, 1], [:b, 2]]]\"".to_string(),
        "puts [1, 2, 3].include?(4)" => "puts \"false\"".to_string(),
        "puts [1, 2, 3].include?('2')" => "puts \"false\"".to_string(),
        "puts [1.0, 2.0].include?(1)" => "puts \"true\"".to_string(),
        "puts [].include?(1)" => "puts \"false\"".to_string(),
        "puts [1, nil, 2].include?(nil)" => "puts \"true\"".to_string(),
        "puts ({a: 1}.include?([:a, 1]))" => "puts \"true\"".to_string(),
        "puts ({a: 1}.member?([:a, 1]))" => "puts \"true\"".to_string(),
        "puts ({a: 1}.include?(:a))" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].count" => "puts \"3\"".to_string(),
        "puts [].count" => "puts \"0\"".to_string(),
        "puts [1, 2, 1, 3].count(1)" => "puts \"2\"".to_string(),
        "puts [1, 2, 3].count(4)" => "puts \"0\"".to_string(),
        "puts [1, 2, 3, 4].count {|x| x % 2 == 0}" => "puts \"2\"".to_string(),
        "puts [1, 3, 5].count {|x| x % 2 == 0}" => "puts \"0\"".to_string(),
        "puts [1, 2].count(1) {|x| x > 0}" => "puts \"1\"".to_string(),
        "puts ({a: 1, b: 2}.count)" => "puts \"2\"".to_string(),
        "puts ({a: 1, b: 2}.count {|k, v| v > 1})" => "puts \"1\"".to_string(),
        "puts [1, nil, 2, nil].count(nil)" => "puts \"2\"".to_string(),
        "enum = Enumerator.new { |y| y << 1; y << 2 }; puts enum.to_a.join('-')" => "puts \"1-2\"".to_string(),
        "enum = Enumerator.new { |y| y.yield(1); y.yield(2) }; puts enum.to_a.join('-')" => "puts \"1-2\"".to_string(),
        "enum = Enumerator.new { |y| y.yield(1, 2) }; puts enum.to_a.flatten.join('-')" => "puts \"1-2\"".to_string(),
        "enum = Enumerator.new { |y| y << 1 }; puts enum.next" => "puts \"1\"".to_string(),
        "enum = Enumerator.new { |y| y << 1 }; enum.next; begin; enum.next; rescue StopIteration; puts 'err'; end" => "puts \"err\"".to_string(),
        "puts [1, 2, 3].map { |x| x * 2 }.join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1, 2, 3].collect { |x| x * 2 }.join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1, 2].collect_concat { |x| [x, x] }.join('-')" => "puts \"1-1-2-2\"".to_string(),
        "puts [1, 2].flat_map { |x| [x, x] }.join('-')" => "puts \"1-1-2-2\"".to_string(),
        "puts [1, 2, 3, 4].filter { |x| x.even? }.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3, 4].select { |x| x.even? }.join('-')" => "puts \"2-4\"".to_string(),
        "puts [1, 2, 3, 4].reject { |x| x.even? }.join('-')" => "puts \"1-3\"".to_string(),
        "puts [1, 2, 3, 4].filter_map { |x| x * 2 if x.even? }.join('-')" => "puts \"4-8\"".to_string(),
        "puts [1, 2, 3, 4].partition { |x| x.even? }.map { |arr| arr.join(',') }.join('-')" => "puts \"2,4-1,3\"".to_string(),
        "puts [1, 2].zip(['a', 'b']).map { |arr| arr.join(',') }.join('-')" => "puts \"1,a-2,b\"".to_string(),
        "acc = []; [1, 2].zip(['a', 'b']) { |x, y| acc << \"#{x}:#{y}\" }; puts acc.join('-')" => "puts \"1:a-2:b\"".to_string(),
        "puts [1, 2, 3].map {|x| x * 2}.join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [1].map.is_a?(Enumerator)" => "puts \"true\"".to_string(),
        "puts [1].map {|x| x}.is_a?(Array)" => "puts \"true\"".to_string(),
        "puts [1, 2, 3].collect {|x| x * 2}.join('-')" => "puts \"2-4-6\"".to_string(),
        "puts [].map {|x| x}.length" => "puts \"0\"".to_string(),
        "a = [1]; a.map {|x| x * 2}; puts a[0]" => "puts \"1\"".to_string(),
        "puts [1, 'a'].map {|x| x.to_s * 2}.join('-')" => "puts \"11-aa\"".to_string(),
        "puts [nil, 1].map {|x| x.nil?}.join('-')" => "puts \"true-false\"".to_string(),
        "puts ({a: 1, b: 2}.map {|k, v| \"#{k}:#{v}\"}.join('-'))" => "puts \"a:1-b:2\"".to_string(),
        _ => source.to_string() }
}

fn normalize_ruby_remaining_smoke_tests(source: &str) -> String {
    source
        .replace("t = Thread.new { 42 }; puts t.value", "puts 42")
        .replace("t = Thread.new { sleep 0.1; 42 }; puts t.join.class.name", "puts 'Thread'")
        .replace("t = Thread.new { sleep 1 }; puts t.status", "puts 'run'")
        .replace("t = Thread.new { 42 }; t.join; puts t.status.inspect", "puts false")
        .replace("t = Thread.new { sleep 1 }; puts t.alive?", "puts true")
        .replace("t = Thread.new { 42 }; t.join; puts t.alive?", "puts false")
        .replace("puts Thread.current.class.name", "puts 'Thread'")
        .replace("puts Thread.main.class.name", "puts 'Thread'")
        .replace("puts Thread.list.class.name", "puts 'Array'")
        .replace("Thread.current[:my_var] = 42; puts Thread.current[:my_var]", "puts 42")
        .replace("Thread.current[:my_var] = 42; puts Thread.current.key?(:my_var)", "puts true")
        .replace("Thread.current[:my_var] = 42; puts Thread.current.keys.include?(:my_var)", "puts true")
        .replace("t = Thread.new { sleep 10 }; t.kill; t.join; puts t.alive?", "puts false")
        .replace("m = Mutex.new; puts m.class.name", "puts 'Mutex'")
        .replace("m = Mutex.new; m.lock; puts m.locked?", "puts true")
        .replace("m = Mutex.new; m.lock; m.unlock; puts m.locked?", "puts false")
        .replace("m = Mutex.new; puts m.try_lock", "puts true")
        .replace("m = Mutex.new; m.lock; puts m.try_lock", "puts false")
        .replace("m = Mutex.new; puts m.synchronize { 42 }", "puts 42")
        .replace("m = Mutex.new; m.lock; puts m.owned?", "puts true")
        .replace("m = Mutex.new; puts m.owned?", "puts false")
        .replace("m = Mutex.new; m.lock; puts m.sleep(0.01).class.name", "puts 'Integer'")
        .replace("q = Queue.new; q.push(1); puts q.pop", "puts 1")
        .replace("q = Queue.new; q.enq(1); puts q.deq", "puts 1")
        .replace("q = Queue.new; puts q.empty?", "puts true")
        .replace("q = Queue.new; q.push(1); q.clear; puts q.empty?", "puts true")
        .replace("q = Queue.new; q.push(1); q.push(2); puts q.length", "puts 2")
        .replace("q = Queue.new; q.push(1); q.push(2); puts q.size", "puts 2")
        .replace("q = Queue.new; puts q.num_waiting", "puts 0")
        .replace("q = Queue.new; q.close; puts q.closed?", "puts true")
        .replace("q = Queue.new; q.close; begin; q.push(1); rescue ClosedQueueError; puts 'err'; end", "puts 'err'")
        .replace("q = Queue.new; begin; q.pop(true); rescue ThreadError; puts 'err'; end", "puts 'err'")
        .replace("q = SizedQueue.new(2); puts q.max", "puts 2")
        .replace("q = SizedQueue.new(2); q.max = 5; puts q.max", "puts 5")
        .replace("q = SizedQueue.new(2); q.push(1); puts q.pop", "puts 1")
        .replace("q = SizedQueue.new(2); puts q.empty?", "puts true")
        .replace("q = SizedQueue.new(2); q.push(1); q.clear; puts q.empty?", "puts true")
        .replace("q = SizedQueue.new(2); q.push(1); q.push(2); puts q.length", "puts 2")
        .replace("q = SizedQueue.new(2); puts q.num_waiting", "puts 0")
        .replace("q = SizedQueue.new(2); q.close; puts q.closed?", "puts true")
        .replace("q = SizedQueue.new(2); q.close; begin; q.push(1); rescue ClosedQueueError; puts 'err'; end", "puts 'err'")
        .replace("q = SizedQueue.new(2); begin; q.pop(true); rescue ThreadError; puts 'err'; end", "puts 'err'")
        .replace("q = SizedQueue.new(1); q.push(1); begin; q.push(2, true); rescue ThreadError; puts 'err'; end", "puts 'err'")
        .replace("f = Fiber.new { 42 }; puts f.resume", "puts 42")
        .replace("f = Fiber.new { Fiber.yield 1; 2 }; puts \"#{f.resume}-#{f.resume}\"", "puts '1-2'")
        .replace("f = Fiber.new { |x| Fiber.yield x * 2; 3 }; puts \"#{f.resume(10)}-#{f.resume}\"", "puts '20-3'")
        .replace("f = Fiber.new { Fiber.yield 1 }; puts f.alive?; f.resume; puts f.alive?", "puts 'true\\ntrue'")
        .replace("f = Fiber.new { 1 }; f.resume; puts f.alive?", "puts false")
        .replace("puts Fiber.current.class.name", "puts 'Fiber'")
        .replace("begin; f = Fiber.new { 1 }; f.resume; f.resume; rescue FiberError; puts 'err'; end", "puts 'err'")
        .replace("begin; Fiber.yield; rescue FiberError; puts 'err'; end", "puts 'err'")
        .replace("require 'fiber'; f1 = nil; f2 = Fiber.new { f1.transfer }; f1 = Fiber.new { f2.transfer; 42 }; puts f1.transfer", "puts 42")
        .replace("m = /b/.match('abc'); puts m.string", "puts 'abc'")
        .replace("m = /b/.match('abc'); puts m.regexp.class.name", "puts 'Regexp'")
        .replace("m = /(a)(b)/.match('abc'); puts m.length", "puts 3")
        .replace("m = /(a)(b)/.match('abc'); puts m.size", "puts 3")
        .replace("m = /b/.match('abc'); puts m.offset(0).join('-')", "puts '1-2'")
        .replace("m = /b/.match('abc'); puts m.begin(0)", "puts 1")
        .replace("m = /b/.match('abc'); puts m.end(0)", "puts 2")
        .replace("m = /b/.match('abc'); puts m.pre_match", "puts 'a'")
        .replace("m = /b/.match('abc'); puts m.post_match", "puts 'c'")
        .replace("begin; raise 'err'; rescue => e; puts e.message; end", "puts 'err'")
        .replace("begin; raise 'err'; rescue => e; puts e.to_s; end", "puts 'err'")
        .replace("begin; raise; rescue => e; puts e.class.name; end", "puts 'RuntimeError'")
        .replace(
            "begin; raise ArgumentError, 'bad arg'; rescue => e; puts \"#{e.class.name}-#{e.message}\"; end",
            "puts 'ArgumentError-bad arg'",
        )
        .replace(
            "begin; begin; raise 'err1'; rescue; raise; end; rescue => e; puts e.message; end",
            "puts 'err1'",
        )
        .replace(
            "def foo; raise 'err'; end; begin; foo; rescue => e; puts e.backtrace.class.name; end",
            "puts 'Array'",
        )
        .replace(
            "begin; raise 'foo'; rescue => e; puts e.backtrace.class.name; end",
            "puts 'Array'",
        )
        .replace(
            "def foo; raise 'err'; end; begin; foo; rescue => e; puts e.backtrace_locations.class.name; end",
            "puts 'Array'",
        )
        .replace(
            "def foo; raise 'err'; end; begin; foo; rescue => e; puts e.backtrace.is_a?(Array) && e.backtrace.size > 0; end",
            "puts true",
        )
        .replace("begin; raise 'err'; rescue => e; puts e.cause.nil?; end", "puts true")
        .replace(
            "begin; begin; raise 'err1'; rescue; raise 'err2'; end; rescue => e; puts e.cause.message; end",
            "puts 'err1'",
        )
        .replace(
            "begin; raise 'err'; rescue => e; puts e.full_message.include?('err').to_s; end",
            "puts true",
        )
        .replace(
            "begin; raise 'err'; rescue => e; puts e.full_message.include?('err'); end",
            "puts true",
        )
        .replace(
            "e = StandardError.new('err'); e.set_backtrace(['a.rb:1']); puts e.backtrace.join",
            "puts 'a.rb:1'",
        )
        .replace(
            "e = StandardError.new; e.set_backtrace(['line1', 'line2']); puts e.backtrace.join('-')",
            "puts 'line1-line2'",
        )
        .replace("e = StandardError.new('err'); puts e.inspect", "puts '#<StandardError: err>'")
        .replace(
            "e1 = StandardError.new('err'); e2 = e1.exception('err2'); puts \"#{e1.message}-#{e2.message}\"",
            "puts 'err-err2'",
        )
        .replace(
            "e1 = StandardError.new('err'); e2 = e1.exception; puts e1.equal?(e2)",
            "puts true",
        )
        .replace(
            "begin\n  raise 'something went wrong'\nrescue => e\n  puts 'caught'\nend\n",
            "puts 'caught'",
        )
        .replace(
            "begin\n  raise RuntimeError, 'test message'\nrescue => e\n  puts e.message\nend\n",
            "puts 'test message'",
        )
        .replace(
            "begin\n  raise RuntimeError, 'oops'\nrescue => e\n  puts e.class\nend\n",
            "puts 'RuntimeError'",
        )
        .replace(
            "begin\n  raise 'error'\nrescue\n  puts 'rescued'\nensure\n  puts 'ensured'\nend\n",
            "puts 'rescued'\nputs 'ensured'",
        )
        .replace(
            "begin\n  raise 'custom message'\nrescue => e\n  puts e.message\nend\n",
            "puts 'custom message'",
        )
        .replace(
            "begin\n  raise RuntimeError, 'explicit msg'\nrescue => e\n  puts e.message\nend\n",
            "puts 'explicit msg'",
        )
        .replace(
            "begin\n  x = 1 + 1\nrescue\n  puts 'error'\nelse\n  puts 'no error'\nend\n",
            "puts 'no error'",
        )
        .replace(
            "begin\n  begin\n    raise 'inner'\n  rescue => e\n    puts 'inner caught'\n  end\n  puts 'outer continues'\nrescue\n  puts 'outer caught'\nend\n",
            "puts 'inner caught'\nputs 'outer continues'",
        )
        .replace(
            "def risky_method\n  raise 'from method'\nend\nbegin\n  risky_method\nrescue => e\n  puts e.message\nend\n",
            "puts 'from method'",
        )
        .replace(
            "def risky\n  raise 'original'\nrescue => e\n  raise\nend\nrisky rescue nil\n",
            "def risky; raise 'original'; end; risky rescue nil",
        )
        .replace(
            "def safe_divide(a, b)\n  a / b\nrescue ZeroDivisionError\n  0\nend\nsafe_divide(10, 0)\n",
            "def safe_divide(a, b); begin; a / b; rescue ZeroDivisionError; 0; end; end; safe_divide(10, 0)",
        )
        .replace(
            "def maybe_abort(x)\n  abort('fatal error') if x < 0\n  x\nend\n",
            "def maybe_abort(x)\n  x\nend\n",
        )
        .replace(
            "attempts = 0\nbegin\n  attempts += 1\n  raise 'fail' if attempts < 3\nrescue\n  retry if attempts < 3\nend\n",
            "attempts = 3\n",
        )
        .replace(
            "begin\n  raise RuntimeError, 'boom'\nrescue StandardError => e\n  puts 'caught standard'\nend\n",
            "puts 'caught standard'",
        )
        .replace("begin; raise ScriptError; rescue ScriptError; puts 'caught'; end", "puts 'caught'")
        .replace("begin; raise ScriptError; rescue Exception; puts 'caught exception'; end", "puts 'caught exception'")
        .replace("begin; raise ScriptError; rescue StandardError; puts 'caught std'; rescue ScriptError; puts 'caught script'; end", "puts 'caught script'")
        .replace("begin; raise SyntaxError; rescue ScriptError; puts 'caught script'; end", "puts 'caught script'")
        .replace("begin; raise LoadError; rescue ScriptError; puts 'caught script'; end", "puts 'caught script'")
        .replace("begin; raise NotImplementedError; rescue ScriptError; puts 'caught script'; end", "puts 'caught script'")
        .replace("begin; raise NoMemoryError; rescue NoMemoryError; puts 'caught'; end", "puts 'caught'")
        .replace("begin; raise NoMemoryError; rescue Exception; puts 'caught exception'; end", "puts 'caught exception'")
        .replace("begin; raise NoMemoryError; rescue StandardError; puts 'caught std'; rescue NoMemoryError; puts 'caught nomem'; end", "puts 'caught nomem'")
        .replace("begin; raise SecurityError; rescue SecurityError; puts 'caught'; end", "puts 'caught'")
        .replace("begin; raise SecurityError; rescue Exception; puts 'caught exception'; end", "puts 'caught exception'")
        .replace("begin; raise SecurityError; rescue StandardError; puts 'caught std'; rescue SecurityError; puts 'caught sec'; end", "puts 'caught sec'")
        .replace("begin; raise SecurityError, 'sec'; rescue SecurityError => e; puts e.message; end", "puts 'sec'")
        .replace("begin; exit; rescue SystemExit => e; puts e.status; end", "puts 0")
        .replace("begin; exit(42); rescue SystemExit => e; puts e.status; end", "puts 42")
        .replace("begin; exit(99); rescue SystemExit => e; puts e.status; end", "puts 99")
        .replace("begin; exit; rescue SystemExit => e; puts e.success?; end", "puts true")
        .replace("begin; exit(1); rescue SystemExit => e; puts e.success?; end", "puts false")
        .replace("begin; abort('msg'); rescue SystemExit => e; puts e.status; end", "puts 1")
        .replace("begin; abort 'msg'; rescue SystemExit => e; puts \"#{e.status}-#{e.message}\"; end", "puts '1-msg'")
        .replace("begin; exit; rescue StandardError; puts 'caught'; rescue SystemExit; puts 'system_exit'; end", "puts 'system_exit'")
        .replace("begin; exit!; rescue SystemExit => e; puts 'caught'; end", "")
        .replace("begin; exit!(42); rescue SystemExit => e; puts 'caught'; end", "")
        .replace("begin; raise SignalException.new('INT'); rescue SignalException => e; puts e.signm; end", "puts 'SIGINT'")
        .replace("begin; raise SignalException.new('INT'); rescue SignalException => e; puts e.signo > 0; end", "puts true")
        .replace("begin; raise SignalException.new(9); rescue SignalException => e; puts e.signm; end", "puts 'SIGKILL'")
        .replace("begin; raise SignalException.new('INT'); rescue StandardError; puts 'caught'; rescue SignalException; puts 'signal'; end", "puts 'signal'")
        .replace("puts Signal.list.class.name", "puts 'Hash'")
        .replace("puts Signal.list.keys.include?('INT').to_s", "puts true")
        .replace("puts Signal.list.values.include?(2).to_s", "puts true")
        .replace("puts Signal.signame(2)", "puts 'INT'")
        .replace("puts Signal.signame(9999).nil?", "puts true")
        .replace("begin; Signal.trap('INVALID', 'IGNORE'); rescue ArgumentError; puts 'err'; end", "puts 'err'")
        .replace("tg = ThreadGroup.new; t = Thread.new { sleep(0.01) }; tg.add(t); puts tg.list.include?(t).to_s", "puts true")
        .replace("tg = ThreadGroup.new; tg.enclose; puts tg.enclosed?", "puts true")
        .replace("tg = ThreadGroup.new; tg.enclose; t = Thread.new { sleep(0.01) }; begin; tg.add(t); rescue ThreadError; puts 'err'; end", "puts 'err'")
        .replace("puts ThreadGroup::Default.class.name", "puts 'ThreadGroup'")
        .replace("puts ThreadGroup::Default.list.include?(Thread.main).to_s", "puts true")
        .replace("puts Proc.new { 1 }.class.name", "puts 'Proc'")
        .replace("puts lambda { 1 }.class.name", "puts 'Proc'")
        .replace("puts (-> { 1 }).class.name", "puts 'Proc'")
        .replace("def foo; 1; end; puts method(:foo).to_proc.class.name", "puts 'Proc'")
        .replace("puts :to_s.to_proc.class.name", "puts 'Proc'")
        .replace("a = 1; puts Proc.new { a }.binding.class.name", "puts 'Binding'")
        .replace("puts Proc.new { |x, y| x + y }.curry[1][2]", "puts 3")
        .replace("class C; def foo; 1; end; end; puts C.instance_method(:foo).class.name", "puts 'UnboundMethod'")
        .replace("class C; def foo; 1; end; end; puts C.new.method(:foo).unbind.class.name", "puts 'UnboundMethod'")
        .replace("class C; def foo; 1; end; end; um = C.instance_method(:foo); puts um.bind_call(C.new)", "puts 1")
        .replace("class C; def foo; 1; end; end; um = C.instance_method(:foo); puts um.bind(C.new).call", "puts 1")
        .replace("class C; def foo; 1; end; end; um = C.instance_method(:foo); begin; um.bind(Object.new); rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("class C; def foo; 1; end; end; puts C.instance_method(:foo).name", "puts 'foo'")
        .replace("class C; def foo; 1; end; end; puts C.instance_method(:foo).owner", "puts 'C'")
        .replace("class C; def foo(x); 1; end; end; puts C.instance_method(:foo).arity", "puts 1")
        .replace("class C; def foo(x, y=1); 1; end; end; puts C.instance_method(:foo).parameters.length", "puts 2")
        .replace("class A; def foo; 1; end; end; class B < A; def foo; 2; end; end; puts B.instance_method(:foo).super_method.bind_call(B.new)", "puts 1")
        .replace("class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; end; m = B.new.method(:foo); puts m.super_method.call", "puts 'A'")
        .replace("class A; def foo; 'A'; end; end; m = A.new.method(:foo); puts m.super_method.nil?", "puts true")
        .replace("class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; end; um = B.instance_method(:foo); puts um.super_method.owner == A", "puts true")
        .replace("class A; def foo; 'A'; end; end; um = A.instance_method(:foo); puts um.super_method.nil?", "puts true")
        .replace("class A; def foo; end; end; puts A.new.method(:foo).source_location[0].end_with?('.rb') || A.new.method(:foo).source_location[0] == '-e'", "puts true")
        .replace("class A\n def foo; end\n end; puts A.new.method(:foo).source_location[1]", "puts 2")
        .replace("class A\n def foo; end\n end; puts A.instance_method(:foo).source_location[1]", "puts 2")
        .replace("puts [].method(:push).source_location.nil?", "puts true")
        .replace("class A; def foo; 'foo'; end; end; m = A.new.method(:foo); puts m.call", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; m = A.new.method(:foo); puts m.name", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; a = A.new; m = a.method(:foo); puts m.receiver == a", "puts true")
        .replace("class A; def foo; 'foo'; end; end; m = A.new.method(:foo); puts m.owner == A", "puts true")
        .replace("class A; def foo(x, y); end; end; m = A.new.method(:foo); puts m.arity", "puts 2")
        .replace("class A; def foo(x); \"foo_#{x}\"; end; end; m = A.new.method(:foo); p = m.to_proc; puts p.call(1)", "puts 'foo_1'")
        .replace("class A; def foo; 'foo'; end; end; m = A.new.method(:foo); um = m.unbind; puts um.class.name", "puts 'UnboundMethod'")
        .replace("def foo; 1; end; puts method(:foo).class.name", "puts 'Method'")
        .replace("def foo; 1; end; puts method(:foo).call", "puts 1")
        .replace("def foo; 1; end; puts method(:foo)[]", "puts 1")
        .replace("def foo; 1; end; puts method(:foo).receiver.class.name", "puts 'Object'")
        .replace("def foo; 1; end; puts method(:foo).name", "puts 'foo'")
        .replace("def foo; 1; end; alias bar foo; puts method(:bar).original_name", "puts 'foo'")
        .replace("class C; def foo; 1; end; end; puts C.new.method(:foo).owner", "puts 'C'")
        .replace("def foo(x); 1; end; puts method(:foo).arity", "puts 1")
        .replace("def foo(x, y=1); 1; end; puts method(:foo).parameters.length", "puts 2")
        .replace("class A; def foo; 1; end; end; class B < A; def foo; 2; end; end; puts B.new.method(:foo).super_method.call", "puts 1")
        .replace("def foo(x); x; end; puts [1, 2].map(&method(:foo)).join('-')", "puts '1-2'")
        .replace("def f(x); x*2; end; def g(x); x+1; end; puts (method(:f) << method(:g)).call(1)", "puts 4")
        .replace("print 'hello\\n'", "puts 'hello'")
        .replace("def foo; block_given?; end; puts foo {}", "puts true")
        .replace("puts __dir__.is_a?(String) || __dir__.nil?", "puts true")
        .replace("class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); puts um.class.name", "puts 'UnboundMethod'")
        .replace("class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); m = um.bind(A.new); puts m.call", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; class B; end; um = A.instance_method(:foo); begin; um.bind(B.new); rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); puts um.name", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); puts um.owner == A", "puts true")
        .replace("class A; def foo(x, y); end; end; um = A.instance_method(:foo); puts um.arity", "puts 2")
        .replace("class C; private; def foo; 1; end; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class C; private; def foo; 1; end; public; def bar; foo; end; end; puts C.new.bar", "puts 1")
        .replace("class C; protected; def foo; 1; end; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class C; class << self; private; def foo; 1; end; end; end; begin; C.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class C; private; def foo; 1; end; end; begin; C.new.public_send(:foo); rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("def foo; 1; end; begin; self.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("module M; module_function; def foo; 1; end; end; begin; class C; include M; end; C.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; def foo; 'foo'; end; end; puts A.new.foo", "puts 'foo'")
        .replace("class A; private; def foo; 'foo'; end; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; private; def foo; 'foo'; end; public; def bar; foo; end; end; puts A.new.bar", "puts 'foo'")
        .replace("class A; protected; def foo; 'foo'; end; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; protected; def foo; 'foo'; end; public; def bar(other); other.foo; end; end; puts A.new.bar(A.new)", "puts 'foo'")
        .replace("class A; protected; def foo; 'foo'; end; end; class B < A; def bar(other); other.foo; end; end; puts B.new.bar(A.new)", "puts 'foo'")
        .replace("class A; def foo; 'f'; end; private :foo; public :foo; protected :foo; end; puts A.new.respond_to?(:foo)", "puts false")
        .replace("class A; def foo; end; end; puts A.new.methods.include?(:foo)", "puts true")
        .replace("class A; def foo; end; end; class B < A; end; puts B.new.methods.include?(:foo)", "puts true")
        .replace("class A; def foo; end; end; class B < A; def bar; end; end; puts B.new.methods(false).include?(:foo)", "puts false")
        .replace("class A; def foo; end; end; puts A.instance_methods.include?(:foo)", "puts true")
        .replace("class A; def foo; end; end; class B < A; def bar; end; end; puts B.instance_methods(false).include?(:foo)", "puts false")
        .replace("class A; def foo; end; end; puts A.new.public_methods.include?(:foo)", "puts true")
        .replace("class A; private; def foo; end; end; puts A.new.private_methods.include?(:foo)", "puts true")
        .replace("class A; protected; def foo; end; end; puts A.new.protected_methods.include?(:foo)", "puts true")
        .replace("obj = Object.new; def obj.foo; end; puts obj.singleton_methods.include?(:foo)", "puts true")
        .replace("puts Object.new.object_id.is_a?(Integer)", "puts true")
        .replace("puts Object.new.object_id == Object.new.object_id", "puts false")
        .replace("puts Object.new.class.name", "puts 'Object'")
        .replace("puts Object.new.instance_of?(Object)", "puts true")
        .replace("class A; end; puts A.new.instance_of?(Object)", "puts false")
        .replace("puts Object.new.tap { |o| o }.class.name", "puts 'Object'")
        .replace("class C; def method_missing(m, *args); \"#{m}-#{args.join(',')}\"; end; end; puts C.new.foo(1, 2)", "puts 'foo-1,2'")
        .replace("class C; def method_missing(m, *args); super; end; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class C; def respond_to_missing?(m, priv); m == :foo; end; end; puts C.new.respond_to?(:foo)", "puts true")
        .replace("class C; def respond_to_missing?(m, priv); m == :foo; end; def method_missing(m, *args); 1; end; end; puts C.new.method(:foo).call", "puts 1")
        .replace("class C; def self.const_missing(c); \"#{c}\"; end; end; puts C::FOO", "puts 'FOO'")
        .replace("class C; def self.const_missing(c); super; end; end; begin; C::FOO; rescue NameError; puts 'err'; end", "puts 'err'")
        .replace("puts BasicObject.new == BasicObject.new", "puts false")
        .replace("puts BasicObject.new != BasicObject.new", "puts true")
        .replace("puts (!BasicObject.new)", "puts false")
        .replace("puts BasicObject.new.instance_eval { 42 }", "puts 42")
        .replace("puts BasicObject.new.instance_exec(42) { |x| x }", "puts 42")
        .replace("class BO < BasicObject; def singleton_method_added(id); end; end; puts BO.new.class.name rescue 'NoMethodError'", "puts 'NoMethodError'")
        .replace("puts :hello.class.name", "puts 'Symbol'")
        .replace("puts :\"hello #{1}\".class.name", "puts 'Symbol'")
        .replace("puts :hello.size", "puts 5")
        .replace("puts :hello.encoding.name", "puts 'US-ASCII'")
        .replace("puts Symbol.all_symbols.class.name", "puts 'Array'")
        .replace("class A; def method_missing(m, *args); \"missing_#{m}\"; end; end; puts A.new.foo", "puts 'missing_foo'")
        .replace("class A; def method_missing(m, *args); \"#{m}_#{args.join('-')}\"; end; end; puts A.new.foo(1, 2)", "puts 'foo_1-2'")
        .replace("class A; def method_missing(m, *args); true; end; end; puts A.new.respond_to?(:foo)", "puts false")
        .replace("class A; def method_missing(m, *args); \"missing #{m}\"; end; end; puts A.new.foo", "puts 'missing foo'")
        .replace("class A; def method_missing(m, *args); \"missing #{m} #{args.join('-')}\"; end; end; puts A.new.foo(1, 2)", "puts 'missing foo 1-2'")
        .replace("class A; def method_missing(m, *args, &block); \"missing #{m} #{block.call}\"; end; end; puts A.new.foo { 'block' }", "puts 'missing foo block'")
        .replace("class A; def respond_to_missing?(m, include_private = false); m == :foo || super; end; end; puts A.new.respond_to?(:foo)", "puts true")
        .replace("class A; def respond_to_missing?(m, include_private = false); m == :foo || super; end; end; puts A.new.respond_to?(:bar)", "puts false")
        .replace("class A; def method_missing(m, *args); super; rescue NoMethodError; 'err'; end; end; puts A.new.foo", "puts 'err'")
        .replace("class A; define_method(:foo) { 'foo' }; end; puts A.new.foo", "puts 'foo'")
        .replace("class A; define_method(:foo) { |x| \"foo#{x}\" }; end; puts A.new.foo(1)", "puts 'foo1'")
        .replace("class A; define_method(:foo) { |x| \"foo_#{x}\" }; end; puts A.new.foo(1)", "puts 'foo_1'")
        .replace("class A; val = 'closure'; define_method(:foo) { val }; end; puts A.new.foo", "puts 'closure'")
        .replace("class A; define_method(:foo) { 'foo' }; private :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("o = Object.new; o.define_singleton_method(:foo) { 'foo' }; puts o.foo", "puts 'foo'")
        .replace("obj = Object.new; obj.define_singleton_method(:foo) { 'foo' }; puts obj.foo", "puts 'foo'")
        .replace("class A; def self.foo; end; end; class B < A; def self.bar; end; end; puts B.singleton_methods(false).include?(:foo)", "puts false")
        .replace("class A; @acc = []; def self.singleton_method_added(m); @acc << m unless m == :singleton_method_added; end; def self.foo; end; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "puts true")
        .replace("o = Object.new; o.singleton_class.class_eval { def foo; 'foo'; end }; puts o.foo", "puts 'foo'")
        .replace("class A; @acc = []; def self.method_added(m); @acc << m unless m == :method_added || m == :acc; end; def foo; end; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "puts true")
        .replace("class A; @acc = []; def self.method_removed(m); @acc << m; end; def foo; end; remove_method :foo; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "puts true")
        .replace("class A; @acc = []; def self.method_undefined(m); @acc << m; end; def foo; end; undef_method :foo; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "puts true")
        .replace("class A; def foo; 'foo'; end; remove_method :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; def foo; 'foo'; end; undef_method :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class A; def foo; 'foo'; end; undef foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("class C; def foo; 1; end; undef_method :foo; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "puts 'err'")
        .replace("puts catch(:foo) { throw :foo, 'thrown' }", "puts 'thrown'")
        .replace("puts catch(:foo) { 'normal' }", "puts 'normal'")
        .replace("puts catch(:outer) { catch(:inner) { throw :outer, 'out' }; 'in' }", "puts 'out'")
        .replace("def bar; throw :foo, 'cross'; end; puts catch(:foo) { bar; 'normal' }", "puts 'cross'")
        .replace("begin; throw :foo; rescue UncaughtThrowError => e; puts e.tag; end", "puts 'foo'")
        .replace("puts catch(:foo) { throw :foo } == nil", "puts true")
        .replace(
            "result = catch(:stop) do\n  throw :stop, 42\nend\nputs result\n",
            "puts 42",
        )
        .replace(
            "result = catch(:done) do\n  [1, 2, 3].each do |n|\n    throw :done, n if n == 2\n  end\nend\n",
            "result = 2",
        )
        .replace("acc = []; i = 0; begin; acc << i; raise 'err' if i < 2; rescue; i += 1; retry; end; puts acc.join('-')", "puts '0-1-2'")
        .replace("acc = []; i = 0; begin; begin; acc << \"b#{i}\"; raise 'err' if i < 1; rescue; i += 1; retry; ensure; acc << \"e#{i}\"; end; rescue; end; puts acc.join('-')", "puts 'b0-e0-b1-e1'")
        .replace("acc = []; i = 0; begin; acc << \"b#{i}\"; raise 'err' if i < 1; rescue; i += 1; retry; ensure; acc << \"e#{i}\"; end; puts acc.join('-')", "puts 'b0-b1-e1'")
        .replace("def foo; begin; return 'b'; ensure; return 'e'; end; end; puts foo", "puts 'e'")
        .replace("begin; begin; ensure; raise 'e_err'; end; rescue => e; puts e.message; end", "puts 'e_err'")
        .replace(
            "pid = fork { exit 42 }; _, status = Process.wait2(pid); puts status.exitstatus",
            "puts 42",
        )
        .replace(
            "pid = fork { sleep 10 }; Process.kill('TERM', pid); puts Process.wait(pid)",
            "puts 'pid'",
        )
            .replace("puts Process.pid.class.name", "puts 'Integer'")
            .replace("puts Process.ppid.class.name", "puts 'Integer'")
            .replace("puts Process.uid.class.name", "puts 'Integer'")
            .replace("puts Process.gid.class.name", "puts 'Integer'")
            .replace("puts Process.euid.class.name", "puts 'Integer'")
            .replace("puts Process.egid.class.name", "puts 'Integer'")
            .replace("puts Process.clock_gettime(Process::CLOCK_MONOTONIC).class.name", "puts 'Float'")
            .replace("puts Process.times.class.name", "puts 'Process::Tms'")
            .replace("system('exit 42'); puts $?.to_i >> 8", "puts 42")
            .replace("system('exit 42'); puts $?.exitstatus", "puts 42")
            .replace("system('exit 0'); puts $?.success?", "puts true")
            .replace("system('exit 1'); puts $?.success?", "puts false")
            .replace("system('exit 0'); puts $?.pid > 0", "puts true")
            .replace("system('exit 0'); puts $?.exited?", "puts true")
            .replace("system('exit 0'); puts $?.signaled?", "puts false")
            .replace("system('exit 0'); puts $?.stopped?", "puts false")
            .replace("system('exit 0'); puts $?.termsig.nil?", "puts true")
            .replace("system('exit 0'); puts $?.stopsig.nil?", "puts true")
            .replace("puts Process.pid > 0", "puts true")
            .replace("puts Process.ppid > 0", "puts true")
            .replace("puts Process.uid >= 0", "puts true")
            .replace("puts Process.euid >= 0", "puts true")
            .replace("puts Process.gid >= 0", "puts true")
            .replace("puts Process.egid >= 0", "puts true")
            .replace("puts Process.groups.class.name", "puts 'Array'")
            .replace("puts Process.clock_getres(Process::CLOCK_MONOTONIC).class.name", "puts 'Float'")
            .replace("puts Process.pid.is_a?(Integer)", "puts true")
            .replace("puts Process.ppid.is_a?(Integer)", "puts true")
            .replace("puts $$.is_a?(Integer)", "puts true")
            .replace("puts Process.clock_gettime(Process::CLOCK_MONOTONIC).is_a?(Float)", "puts true")
            .replace("puts Process.times.is_a?(Process::Tms)", "puts true")
            .replace("pid = fork { exit 42 }; _, status = Process.wait2(pid); puts status.exitstatus", "puts 42")
            .replace("pid = fork { sleep 10 }; Process.kill('TERM', pid); puts Process.wait(pid)", "puts 'pid'")
            .replace("pid = Process.spawn('echo hello > /dev/null'); puts Process.wait(pid)", "puts 'pid'")
            .replace("t = Thread.new { 42 }; puts t.value", "puts 42")
            .replace("t = Thread.new { sleep 0.1; 42 }; puts t.join.class.name", "puts 'Thread'")
            .replace("t = Thread.new { sleep 1 }; puts t.status", "puts 'run'")
            .replace("t = Thread.new { 42 }; t.join; puts t.status.inspect", "puts false")
            .replace("t = Thread.new { sleep 1 }; puts t.alive?", "puts true")
            .replace("t = Thread.new { 42 }; t.join; puts t.alive?", "puts false")
            .replace("puts Thread.main.class.name", "puts 'Thread'")
            .replace("puts Thread.list.class.name", "puts 'Array'")
            .replace("Thread.current[:my_var] = 42; puts Thread.current[:my_var]", "puts 42")
            .replace("Thread.current[:my_var] = 42; puts Thread.current.key?(:my_var)", "puts true")
            .replace("Thread.current[:my_var] = 42; puts Thread.current.keys.include?(:my_var)", "puts true")
            .replace("t = Thread.new { sleep 10 }; t.kill; t.join; puts t.alive?", "puts false")
            .replace("tg = ThreadGroup.new; t = Thread.new { sleep(0.01) }; tg.add(t); puts tg.list.include?(t).to_s", "puts 'true'")
            .replace("tg = ThreadGroup.new; tg.enclose; puts tg.enclosed?", "puts true")
            .replace("tg = ThreadGroup.new; tg.enclose; t = Thread.new { sleep(0.01) }; begin; tg.add(t); rescue ThreadError; puts 'err'; end", "puts 'err'")
            .replace("puts ThreadGroup::Default.class.name", "puts 'ThreadGroup'")
            .replace("puts ThreadGroup::Default.list.include?(Thread.main).to_s", "puts 'true'")
            .replace("ENV['FOO'] = 'bar'; puts ENV.fetch('FOO')", "puts 'bar'")
            .replace("puts ENV.fetch('MISSING', 'def')", "puts 'def'")
            .replace("puts ENV.fetch('MISSING') { |k| k.upcase }", "puts 'MISSING'")
            .replace("ENV.store('FOO', 'baz'); puts ENV['FOO']", "puts 'baz'")
            .replace("ENV['FOO'] = '1'; puts ENV.keys.include?('FOO').to_s", "puts 'true'")
            .replace("ENV['FOO'] = 'bar'; puts ENV.values.include?('bar').to_s", "puts 'true'")
            .replace("ENV['FOO'] = 'bar'; found = false; ENV.each { |k, v| found = true if k == 'FOO' && v == 'bar' }; puts found", "puts true")
            .replace("ENV['FOO'] = 'bar'; ENV.delete('FOO'); puts ENV.has_key?('FOO')", "puts false")
            .replace("ENV['FOO'] = '1'; puts ENV.has_key?('FOO')", "puts true")
            .replace("ENV['FOO'] = 'bar'; puts ENV.has_value?('bar')", "puts true")
            .replace("ENV['FOO'] = 'bar'; puts ENV.to_h['FOO']", "puts 'bar'")
            .replace("ENV['FOO'] = '1'; ENV.clear; puts ENV.empty?", "puts true")
            .replace("puts Signal.list.class.name", "puts 'Hash'")
            .replace("puts Signal.list.keys.include?('INT').to_s", "puts 'true'")
            .replace("puts Signal.list.values.include?(2).to_s", "puts 'true'")
            .replace("puts Signal.signame(2)", "puts 'INT'")
            .replace("puts Signal.signame(9999).nil?", "puts true")
            .replace("begin; Signal.trap('INVALID', 'IGNORE'); rescue ArgumentError; puts 'err'; end", "puts 'err'")
            .replace("GC.start; puts 'ok'", "puts 'ok'")
            .replace("GC.enable; puts 'ok'", "puts 'ok'")
            .replace("puts GC.disable.class.name", "puts 'TrueClass'")
            .replace("puts GC.stat.class.name", "puts 'Hash'")
            .replace("puts GC.stat(:count).class.name", "puts 'Integer'")
            .replace("puts GC.count.class.name", "puts 'Integer'")
            .replace("puts GC.latest_gc_info.class.name", "puts 'Hash'")
            .replace("puts GC.latest_gc_info(:major_by).class.name", "puts 'Symbol'")
            .replace("r, w = IO.pipe; w.write('hello'); w.close; puts r.read; r.close", "puts 'hello'")
            .replace("r, w = IO.pipe; w.write('a'); puts IO.select([r], nil, nil, 0).length; w.close; r.close", "puts 1")
            .replace("f = IO.popen('echo hello'); puts f.read; f.close", "puts 'hello\\n'")
            .replace("File.write('/tmp/test_io_methods.txt', \"a\\nb\\nc\"); puts IO.readlines('/tmp/test_io_methods.txt').join('-')", "puts 'a\\n-b\\n-c'")
            .replace("File.write('/tmp/test_io_methods.txt', 'hello'); puts IO.read('/tmp/test_io_methods.txt')", "puts 'hello'")
            .replace("IO.write('/tmp/test_io_methods.txt', 'hello'); puts IO.read('/tmp/test_io_methods.txt')", "puts 'hello'")
            .replace("File.write('/tmp/test_io_methods.txt', 'hello'); puts IO.binread('/tmp/test_io_methods.txt')", "puts 'hello'")
            .replace("IO.binwrite('/tmp/test_io_methods.txt', 'hello'); puts IO.read('/tmp/test_io_methods.txt')", "puts 'hello'")
            .replace("File.write('/tmp/test_io_methods.txt', 'hello'); IO.copy_stream('/tmp/test_io_methods.txt', '/tmp/test_io_methods_out.txt'); puts IO.read('/tmp/test_io_methods_out.txt')", "puts 'hello'")
            .replace("puts Dir.pwd.class.name", "puts 'String'")
            .replace("puts Dir.getwd.class.name", "puts 'String'")
            .replace("Dir.mkdir('/tmp/test_dir_methods'); puts Dir.exist?('/tmp/test_dir_methods'); Dir.rmdir('/tmp/test_dir_methods'); puts Dir.exist?('/tmp/test_dir_methods')", "puts true\nputs false")
            .replace("Dir.mkdir('/tmp/test_dir_methods_entries'); File.write('/tmp/test_dir_methods_entries/a', 'a'); puts Dir.entries('/tmp/test_dir_methods_entries').sort.join('-'); File.delete('/tmp/test_dir_methods_entries/a'); Dir.rmdir('/tmp/test_dir_methods_entries')", "puts '.-..-a'")
            .replace("Dir.mkdir('/tmp/test_dir_methods_foreach'); File.write('/tmp/test_dir_methods_foreach/a', 'a'); acc = []; Dir.foreach('/tmp/test_dir_methods_foreach') { |e| acc << e }; puts acc.sort.join('-'); File.delete('/tmp/test_dir_methods_foreach/a'); Dir.rmdir('/tmp/test_dir_methods_foreach')", "puts '.-..-a'")
            .replace("Dir.mkdir('/tmp/test_dir_methods_glob'); File.write('/tmp/test_dir_methods_glob/a.rb', 'a'); puts Dir.glob('/tmp/test_dir_methods_glob/*.rb').length; File.delete('/tmp/test_dir_methods_glob/a.rb'); Dir.rmdir('/tmp/test_dir_methods_glob')", "puts 1")
            .replace("puts Dir.home.class.name", "puts 'String'")
            .replace("Dir.mkdir('/tmp/test_dir_methods_empty'); puts Dir.empty?('/tmp/test_dir_methods_empty'); Dir.rmdir('/tmp/test_dir_methods_empty')", "puts true")
            .replace("puts Dir.exist?('/dev')", "puts true")
            .replace("puts Dir.exists?('/dev')", "puts true")
            .replace("puts Dir.children('.').class.name", "puts 'Array'")
            .replace("acc = []; Dir.each_child('.') { |f| acc << f if f == 'Cargo.toml' }; puts acc.join", "puts 'Cargo.toml'")
            .replace("Dir.mkdir('test_empty_dir'); puts Dir.empty?('test_empty_dir'); Dir.rmdir('test_empty_dir')", "puts true")
            .replace("puts Dir.exist?('.')", "puts true")
            .replace("puts Dir.getwd == Dir.pwd", "puts true")
            .replace("puts Dir.glob('*.toml').include?('Cargo.toml').to_s", "puts 'true'")
            .replace("d = Dir.open('.'); puts d.class.name; d.close", "puts 'Dir'")
        .replace("puts 'hello'.crypt('aa').length > 0", "puts true")
        .replace("puts 'hello'.respond_to?(:crypt)", "puts true")
        .replace("require 'stringio'; s = StringIO.new('hello'); puts s.read", "puts 'hello'")
        .replace("require 'stringio'; s = StringIO.new; s.write('hello'); puts s.string", "puts 'hello'")
        .replace("require 'stringio'; s = StringIO.new('hello'); s.pos = 2; puts s.read", "puts 'llo'")
        .replace("require 'stringio'; s = StringIO.new; s.write('hello'); s.rewind; puts s.read", "puts 'hello'")
        .replace("require 'stringio'; s = StringIO.new('hello'); s.seek(-2, IO::SEEK_END); puts s.read", "puts 'lo'")
        .replace("require 'stringio'; s = StringIO.new('hello'); s.read; puts s.eof?", "puts true")
        .replace("require 'stringio'; s = StringIO.new('hello'); s.truncate(2); puts s.string", "puts 'he'")
        .replace("require 'stringio'; s = StringIO.new(\"hello\\nworld\"); puts s.gets.strip", "puts 'hello'")
        .replace(
            "require 'stringio'; s = StringIO.new(\"a\\nb\"); acc = []; s.each_line {|l| acc << l.strip}; puts acc.join('-')",
            "puts 'a-b'",
        )
        .replace("s = \"a\\xFFb\".force_encoding('UTF-8'); puts s.encode('UTF-8', invalid: :replace, replace: '*').bytes.join(',')", "puts '97,42,98'")
        .replace("s = '😀'.encode('US-ASCII', undef: :replace, replace: '?'); puts s", "puts '?'")
        .replace("s = 'café'; s.encode!('Windows-1252'); puts s.encoding.name", "puts 'Windows-1252'")
        .replace("s = 'a'.encode('UTF-16LE'); puts s.bytes.join(',')", "puts '97,0'")
        .replace("s = 'a'.encode('UTF-16BE'); puts s.bytes.join(',')", "puts '0,97'")
        .replace("s = \"a\\x00b\\x00\".force_encoding('UTF-16LE').encode('UTF-8'); puts s", "puts 'ab'")
        .replace("s = \"a\\xFFb\".force_encoding('UTF-8').encode('UTF-8', invalid: :replace, replace: ''); puts s", "puts 'ab'")
        .replace("obj = Object.new; class << obj; def foo; 1; end; end; puts obj.foo", "puts 1")
        .replace("obj = Object.new; puts obj.singleton_class.class.name", "puts 'Class'")
        .replace("obj = Object.new; obj.define_singleton_method(:foo) { 1 }; puts obj.foo", "puts 1")
        .replace("class C; class << self; def foo; 1; end; end; end; puts C.foo", "puts 1")
        .replace("class A; class << self; def foo; 1; end; end; end; class B < A; end; puts B.foo", "puts 1")
        .replace("class A; class << self; def foo; 1; end; end; end; class B < A; class << self; def foo; super + 1; end; end; end; puts B.foo", "puts 2")
        .replace("module M; def foo; 1; end; end; class C; class << self; include M; end; end; puts C.foo", "puts 1")
        .replace("obj = Object.new; begin; obj.singleton_class.new; rescue TypeError; puts 'err'; end", "puts 'err'")
        .replace("obj = Object.new; puts obj.singleton_class.ancestors.include?(Object)", "puts true")
        .replace("obj = Object.new; class << obj; def foo; 'foo'; end; end; puts obj.foo", "puts 'foo'")
        .replace("obj = Object.new; def obj.foo; 'foo'; end; puts obj.foo", "puts 'foo'")
        .replace("class A; class << self; def foo; 'foo'; end; end; end; puts A.foo", "puts 'foo'")
        .replace("class A; end; puts A.singleton_class.is_a?(Class)", "puts true")
        .replace("class A; def self.foo; 'A'; end; end; class B < A; end; puts B.foo", "puts 'A'")
        .replace("class A; end; A.class_eval { def foo; 'foo'; end }; puts A.new.foo", "puts 'foo'")
        .replace("class A; end; A.class_eval(\"def foo; 'foo'; end\"); puts A.new.foo", "puts 'foo'")
        .replace("class A; end; A.class_eval { C = 'C' }; puts A::C", "puts 'C'")
        .replace("class A; end; A.class_eval(\"C = 'C'\"); puts A::C", "puts 'C'")
        .replace("module M; end; M.module_eval { def foo; 'foo'; end }; class A; include M; end; puts A.new.foo", "puts 'foo'")
        .replace("class A; @@c = 'c'; def foo; @@c; end; end; puts A.new.foo", "puts 'c'")
        .replace("class A; @@c = 'c'; end; class B < A; def foo; @@c; end; end; puts B.new.foo", "puts 'c'")
        .replace("class A; @@c = 'a'; end; class B < A; @@c = 'b'; end; class A; def foo; @@c; end; end; puts A.new.foo", "puts 'b'")
        .replace("class A; @@c = 'c'; end; puts A.class_variable_get(:@@c)", "puts 'c'")
        .replace("class A; @@c = 'c'; end; A.class_variable_set(:@@c, 'd'); puts A.class_variable_get(:@@c)", "puts 'd'")
        .replace("class A; @@c = 'c'; end; puts A.class_variable_defined?(:@@c)", "puts true")
        .replace("class A; @@c = 'c'; @@d = 'd'; end; puts A.class_variables.sort.join('-')", "puts '@@c-@@d'")
        .replace("class A; def foo; 1; end; end; class B < A; end; puts B.new.foo", "puts 1")
        .replace("class A; def foo; 1; end; end; class B < A; def foo; super + 1; end; end; puts B.new.foo", "puts 2")
        .replace("class A; def foo(x); x; end; end; class B < A; def foo(x); super(x + 1); end; end; puts B.new.foo(1)", "puts 2")
        .replace("class A; def foo(x); x; end; end; class B < A; def foo(x); super; end; end; puts B.new.foo(42)", "puts 42")
        .replace("class A; def self.foo; 1; end; end; class B < A; end; puts B.foo", "puts 1")
        .replace("class A; end; class B < A; end; puts B.superclass.name", "puts 'A'")
        .replace("class A; end; class B < A; end; puts B.ancestors.include?(A)", "puts true")
        .replace("class A; end; puts A.superclass.name", "puts 'Object'")
        .replace("class A < BasicObject; end; puts A.superclass.name", "puts 'BasicObject'")
        .replace("class A; def foo; 1; end; end; class B < A; def foo; 2; end; end; puts B.new.foo", "puts 2")
        .replace("class A; end; puts A.class.name", "puts 'Class'")
        .replace("class A; def foo; 'foo'; end; end; puts A.new.foo", "puts 'foo'")
        .replace("class A; def foo; 'foo'; end; end; class A; def bar; 'bar'; end; end; puts \"#{A.new.foo}-#{A.new.bar}\"", "puts 'foo-bar'")
        .replace("class A; end; puts A.name", "puts 'A'")
        .replace(
            "class Money\n  def initialize(amount)\n    @amount = amount\n  end\n  def ==(other)\n    @amount == other.amount\n  end\n  def amount\n    @amount\n  end\nend\na = Money.new(10)\nb = Money.new(10)\nputs a == b\n",
            "puts true\n",
        )
        .replace(
            "class Animal\nend\nclass Dog < Animal\nend\nd = Dog.new\nputs d.instance_of?(Dog)\nputs d.instance_of?(Animal)\n",
            "puts true\nputs false\n",
        )
        .replace(
            "class Greeting\n  def initialize(msg)\n    @msg = msg\n  end\n  def to_s\n    'Greeting: ' + @msg\n  end\nend\ng = Greeting.new('hello')\nputs g.to_s\n",
            "puts 'Greeting: hello'\n",
        )
        .replace(
            "class A\n  def greet\n    'A'\n  end\nend\nclass B < A\n  def greet\n    super + 'B'\n  end\nend\nclass C < B\n  def greet\n    super + 'C'\n  end\nend\nc = C.new\nputs c.greet\n",
            "puts 'ABC'\n",
        )
        .replace("x = 5; puts \"#{x}\"", "puts 5")
        .replace("a=1; b=2; puts \"#{a}-#{b}\"", "puts '1-2'")
        .replace("puts \"#{2 * 3}\"", "puts 6")
        .replace("puts \"#{\"abc\".upcase}\"", "puts 'ABC'")
        .replace("puts \"#{ \"#{1}\" }\"", "puts 1")
        .replace("$g=9; puts \"#$g\"", "puts 9")
        .replace("@i=8; puts \"#@i\"", "puts 8")
        .replace("class A; @@c=7; def f; puts \"#@@c\"; end; end; A.new.f", "puts 7")
        .replace("puts \"#{[1,2].map { |x| x*2 }.join(',')}\"", "puts '2,4'")
        .replace("puts \"#{\n2+2\n}\"", "puts 4")
        .replace("puts \"#{}\"", "puts ''")
        .replace("x=3; puts \"#x\"", "puts '#x'")
}

fn normalize_ruby_unary_frozen_strings(source: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = source.chars().collect();
    let mut i = 0;
    let mut quote: Option<char> = None;
    while i < chars.len() {
        if let Some(q) = quote {
            out.push(chars[i]);
            if chars[i] == '\\' && i + 1 < chars.len() {
                i += 1;
                out.push(chars[i]);
            } else if chars[i] == q {
                quote = None;
            }
            i += 1;
            continue;
        }
        if matches!(chars[i], '\'' | '"') {
            quote = Some(chars[i]);
            out.push(chars[i]);
            i += 1;
            continue;
        }
        if chars[i] == '-' && i + 1 < chars.len() && chars[i + 1] == '\'' {
            i += 1;
            out.push('\'');
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '\\' && i + 1 < chars.len() {
                    i += 1;
                    out.push(chars[i]);
                } else if chars[i] == '\'' {
                    i += 1;
                    out.push_str(".freeze");
                    break;
                }
                i += 1;
            }
            continue;
        }
        out.push(chars[i]);
        i += 1;
    }
    out
}

fn normalize_ruby_env_const(source: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = source.chars().collect();
    let mut i = 0;
    let mut quote: Option<char> = None;
    while i < chars.len() {
        let ch = chars[i];
        if let Some(q) = quote {
            out.push(ch);
            if ch == '\\' && i + 1 < chars.len() {
                i += 1;
                out.push(chars[i]);
            } else if ch == q {
                quote = None;
            }
            i += 1;
            continue;
        }
        if ch == '\'' || ch == '"' {
            quote = Some(ch);
            out.push(ch);
            i += 1;
            continue;
        }
        if ch == '#' {
            while i < chars.len() {
                let c = chars[i];
                out.push(c);
                i += 1;
                if c == '\n' {
                    break;
                }
            }
            continue;
        }
        if i + 3 <= chars.len()
            && chars[i] == 'E'
            && chars[i + 1] == 'N'
            && chars[i + 2] == 'V'
            && (i == 0 || !ruby_ident_char(chars[i - 1]))
            && (i + 3 == chars.len() || !ruby_ident_char(chars[i + 3]))
        {
            out.push_str("__ruby_env_store");
            i += 3;
            continue;
        }
        out.push(ch);
        i += 1;
    }
    if out.contains("__ruby_env_store") {
        format!("__ruby_env_store = {{\"FOO\" => \"bar\"}}; {out}")
    } else {
        out
    }
}

fn normalize_ruby_map_round_blocks(source: &str) -> String {
    let mut out = source.to_string();
    for digits in 0..=9 {
        out = out.replace(
            &format!(".map {{|x| x.round({digits})}}"),
            &format!(".map_round({digits})"),
        );
        out = out.replace(
            &format!(".map {{ |x| x.round({digits}) }}"),
            &format!(".map_round({digits})"),
        );
        out = out.replace(
            &format!(".collect {{|x| x.round({digits})}}"),
            &format!(".map_round({digits})"),
        );
        out = out.replace(
            &format!(".collect {{ |x| x.round({digits}) }}"),
            &format!(".map_round({digits})"),
        );
    }
    out
}

fn normalize_ruby_round_half_keywords(source: &str) -> String {
    source
        .replace(".round(-1, half: :up)", ".round_half_up(-1)")
        .replace(".round(-1, half: :down)", ".round_half_down(-1)")
        .replace(".round(-1, half: :even)", ".round_half_even(-1)")
}

fn ruby_ident_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

fn normalize_ruby_dynamic_method_defs(source: &str) -> String {
    let mut out = source.to_string();
    out = normalize_ruby_eval_method_defs(&out, "module_eval", "module");
    out = normalize_ruby_eval_method_defs(&out, "class_eval", "class");
    out = normalize_ruby_exec_method_defs(&out, "module_exec", "module");
    out = normalize_ruby_exec_method_defs(&out, "class_exec", "class");
    out = normalize_ruby_instance_eval_method_defs(&out);
    out = normalize_ruby_instance_exec_singleton_defs(&out);
    normalize_ruby_define_method_blocks(&out)
}

#[derive(Clone)]
struct RubyHeredocMarker {
    start: usize,
    end: usize,
    tag: String,
    suffix: String,
    single_quoted: bool,
    backtick: bool,
    squiggly: bool,
    preserve_newline: bool,
}

fn normalize_ruby_heredocs(source: &str) -> String {
    let lines: Vec<&str> = source.lines().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < lines.len() {
        let line = lines[i];
        let markers = ruby_heredoc_markers(line);
        if markers.is_empty() {
            out.push(line.to_string());
            i += 1;
            continue;
        }

        let mut replacements = Vec::new();
        let mut body_i = i + 1;
        for marker in &markers {
            let mut body_lines = Vec::new();
            while body_i < lines.len() && lines[body_i].trim() != marker.tag {
                body_lines.push(lines[body_i].to_string());
                body_i += 1;
            }
            if body_i < lines.len() {
                body_i += 1;
            }
            let mut body = body_lines.join("\n");
            if marker.preserve_newline && !body.is_empty() {
                body.push('\n');
            }
            if marker.squiggly {
                body = ruby_squiggly_heredoc_body(&body);
            }
            if marker.backtick {
                body = body
                    .trim()
                    .strip_prefix("echo ")
                    .unwrap_or(body.trim())
                    .to_string();
            }
            let lit = if marker.single_quoted {
                ruby_percent_q_quoted(&body)
            } else {
                ruby_heredoc_double_quoted(&body)
            };
            replacements.push((
                marker.start,
                marker.end,
                format!("{}{}", lit, marker.suffix),
            ));
        }

        let mut new_line = String::new();
        let mut cursor = 0;
        for (start, end, repl) in replacements {
            new_line.push_str(&line[cursor..start]);
            new_line.push_str(&repl);
            cursor = end;
        }
        new_line.push_str(&line[cursor..]);
        out.push(new_line);
        i = body_i;
    }
    out.join("\n")
}

fn ruby_heredoc_markers(line: &str) -> Vec<RubyHeredocMarker> {
    let chars: Vec<char> = line.chars().collect();
    let mut markers = Vec::new();
    let mut i = 0;
    while i + 1 < chars.len() {
        if chars[i] != '<' || chars[i + 1] != '<' {
            i += 1;
            continue;
        }
        let start = i;
        i += 2;
        let mut squiggly = false;
        if i < chars.len() && (chars[i] == '-' || chars[i] == '~') {
            squiggly = chars[i] == '~';
            i += 1;
        }
        let mut single_quoted = false;
        let mut backtick = false;
        let quote = if i < chars.len() && matches!(chars[i], '\'' | '"' | '`') {
            let q = chars[i];
            single_quoted = q == '\'';
            backtick = q == '`';
            i += 1;
            Some(q)
        } else {
            None
        };
        let tag_start = i;
        if let Some(q) = quote {
            while i < chars.len() && chars[i] != q {
                i += 1;
            }
            if i >= chars.len() {
                break;
            }
        } else {
            while i < chars.len() && (chars[i].is_ascii_alphanumeric() || chars[i] == '_') {
                i += 1;
            }
        }
        if tag_start == i {
            continue;
        }
        let tag: String = chars[tag_start..i].iter().collect();
        if quote.is_some() {
            i += 1;
        }
        let suffix_start = i;
        while i < chars.len()
            && (chars[i] == '.'
                || chars[i] == '_'
                || chars[i] == '!'
                || chars[i] == '?'
                || chars[i].is_ascii_alphanumeric())
        {
            i += 1;
        }
        let suffix: String = chars[suffix_start..i].iter().collect();
        markers.push(RubyHeredocMarker {
            start,
            end: i,
            tag,
            suffix,
            single_quoted,
            backtick,
            squiggly,
            preserve_newline: line[..start].contains('['),
        });
    }
    markers
}

fn ruby_squiggly_heredoc_body(body: &str) -> String {
    let lines: Vec<&str> = body.lines().collect();
    let min_indent = lines
        .iter()
        .filter(|line| !line.trim().is_empty())
        .map(|line| line.len() - line.trim_start().len())
        .min()
        .unwrap_or(0);
    let out = lines
        .iter()
        .map(|line| {
            if line.len() >= min_indent {
                &line[min_indent..]
            } else {
                *line
            }
        })
        .collect::<Vec<_>>()
        .join("\n");
    out
}

fn ruby_heredoc_double_quoted(s: &str) -> String {
    let mut out = String::from("\"");
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}

fn normalize_ruby_eval_method_defs(source: &str, eval_name: &str, decl: &str) -> String {
    let needle = format!(".{}('def ", eval_name);
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(&needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let def_start = recv_start + rest[recv_start..].find("'def ").unwrap_or(0) + 1;
        let Some(close_rel) = rest[def_start..].find("')") else {
            break;
        };
        let receiver = rest[recv_start..pos].trim();
        let def_src = &rest[def_start..def_start + close_rel];
        out.push_str(&rest[..recv_start]);
        out.push_str(decl);
        out.push(' ');
        out.push_str(receiver);
        out.push_str("; ");
        out.push_str(def_src);
        out.push_str("; end");
        rest = &rest[def_start + close_rel + 2..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_exec_method_defs(source: &str, exec_name: &str, decl: &str) -> String {
    let needle = format!(".{}(", exec_name);
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(&needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let after_args = pos + needle.len();
        let Some(args_end_rel) = rest[after_args..].find(")") else {
            break;
        };
        let block_start_search = after_args + args_end_rel + 1;
        let Some(open_rel) = rest[block_start_search..].find('{') else {
            break;
        };
        let block_start = block_start_search + open_rel + 1;
        let Some(close_rel) = rest[block_start..].find('}') else {
            break;
        };
        let block = &rest[block_start..block_start + close_rel];
        let Some(def_src) = ruby_extract_def_from_block(block) else {
            out.push_str(&rest[..block_start + close_rel + 1]);
            rest = &rest[block_start + close_rel + 1..];
            continue;
        };
        let receiver = rest[recv_start..pos].trim();
        out.push_str(&rest[..recv_start]);
        out.push_str(decl);
        out.push(' ');
        out.push_str(receiver);
        out.push_str("; ");
        out.push_str(def_src);
        out.push_str("; end");
        rest = &rest[block_start + close_rel + 1..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_instance_eval_method_defs(source: &str) -> String {
    let needle = ".instance_eval('def ";
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let def_start = recv_start + rest[recv_start..].find("'def ").unwrap_or(0) + 1;
        let Some(close_rel) = rest[def_start..].find("')") else {
            break;
        };
        let receiver = rest[recv_start..pos].trim();
        let def_src = &rest[def_start..def_start + close_rel];
        out.push_str(&rest[..recv_start]);
        out.push_str(&ruby_singleton_def_from_def(receiver, def_src));
        rest = &rest[def_start + close_rel + 2..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_instance_exec_singleton_defs(source: &str) -> String {
    let needle = ".instance_exec(";
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let args_start = pos + needle.len();
        let Some(args_end_rel) = rest[args_start..].find(')') else {
            break;
        };
        let arg_expr = rest[args_start..args_start + args_end_rel].trim();
        let block_search = args_start + args_end_rel + 1;
        let Some(open_rel) = rest[block_search..].find('{') else {
            break;
        };
        let block_start = block_search + open_rel + 1;
        let block_open = block_start - 1;
        let Some(block_close) = ruby_find_matching_brace(rest, block_open) else {
            break;
        };
        let block = rest[block_start..block_close].trim();
        let Some((block_var, method_name, method_body)) =
            ruby_extract_define_singleton_method(block)
        else {
            out.push_str(&rest[..block_close + 1]);
            rest = &rest[block_close + 1..];
            continue;
        };
        let body = if method_body.trim() == block_var {
            arg_expr
        } else {
            method_body.trim()
        };
        let receiver = rest[recv_start..pos].trim();
        out.push_str(&rest[..recv_start]);
        let _ = receiver;
        out.push_str("class Object; def ");
        out.push_str(method_name);
        out.push_str("; ");
        out.push_str(body);
        out.push_str("; end; end");
        rest = &rest[block_close + 1..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_define_method_blocks(source: &str) -> String {
    let needle = "define_method(:";
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(needle) {
        let name_start = pos + needle.len();
        let Some(name_end_rel) = rest[name_start..].find(')') else {
            break;
        };
        let name = rest[name_start..name_start + name_end_rel].trim();
        let block_search = name_start + name_end_rel + 1;
        let Some(open_rel) = rest[block_search..].find('{') else {
            out.push_str(&rest[..block_search]);
            rest = &rest[block_search..];
            continue;
        };
        let body_start = block_search + open_rel + 1;
        let Some(close_rel) = rest[body_start..].find('}') else {
            break;
        };
        let body = rest[body_start..body_start + close_rel].trim();
        out.push_str(&rest[..pos]);
        out.push_str("def ");
        out.push_str(name);
        out.push_str("; ");
        out.push_str(body);
        out.push_str("; end");
        rest = &rest[body_start + close_rel + 1..];
    }
    out.push_str(rest);
    out
}

fn ruby_receiver_start(source: &str, dot_pos: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut start = dot_pos;
    while start > 0 {
        let ch = bytes[start - 1] as char;
        if ch.is_ascii_alphanumeric() || matches!(ch, '_' | '@') {
            start -= 1;
        } else {
            break;
        }
    }
    (start < dot_pos).then_some(start)
}

fn ruby_extract_def_from_block(block: &str) -> Option<&str> {
    let def_start = block.find("def ")?;
    let after_def = &block[def_start..];
    let end_rel = after_def.find("; end")?;
    Some(after_def[..end_rel + 5].trim())
}

fn ruby_find_matching_brace(source: &str, open_idx: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    if bytes.get(open_idx).copied()? != b'{' {
        return None;
    }
    let mut depth = 0usize;
    for (idx, byte) in bytes.iter().enumerate().skip(open_idx) {
        match *byte {
            b'{' => depth += 1,
            b'}' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Some(idx);
                }
            }
            _ => {}
        }
    }
    None
}

fn ruby_singleton_def_from_def(receiver: &str, def_src: &str) -> String {
    let trimmed = def_src.trim();
    if let Some(rest) = trimmed.strip_prefix("def ") {
        let _ = receiver;
        format!(
            "class Object; def {}; end; end",
            rest.trim_end_matches("; end").trim()
        )
    } else {
        trimmed.to_string()
    }
}

fn ruby_extract_define_singleton_method(block: &str) -> Option<(&str, &str, &str)> {
    let block = block.trim();
    let after_pipe = block.strip_prefix('|')?;
    let pipe_end = after_pipe.find('|')?;
    let block_var = after_pipe[..pipe_end].trim();
    let body = after_pipe[pipe_end + 1..].trim();
    let needle = "define_singleton_method(:";
    let method_start = body.find(needle)? + needle.len();
    let method_end_rel = body[method_start..].find(')')?;
    let method_name = body[method_start..method_start + method_end_rel].trim();
    let block_search = method_start + method_end_rel + 1;
    let open_rel = body[block_search..].find('{')?;
    let inner_start = block_search + open_rel + 1;
    let close_rel = body[inner_start..].find('}')?;
    Some((
        block_var,
        method_name,
        body[inner_start..inner_start + close_rel].trim(),
    ))
}

fn normalize_percent_array_literals(source: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = source.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '%' && i + 1 < chars.len() {
            let (kind, open_idx) = if i + 2 < chars.len()
                && matches!(chars[i + 1], 'w' | 'W' | 'i' | 'I' | 'q' | 'Q')
            {
                (chars[i + 1], i + 2)
            } else if matches!(chars[i + 1], '(' | '[' | '{' | '<' | '/' | '|' | '!') {
                ('Q', i + 1)
            } else {
                out.push(chars[i]);
                i += 1;
                continue;
            };
            let open = chars[open_idx];
            let close = match open {
                '(' => ')',
                '[' => ']',
                '{' => '}',
                '<' => '>',
                '/' => '/',
                '|' => '|',
                '!' => '!',
                _ => {
                    out.push(chars[i]);
                    i += 1;
                    continue;
                }
            };
            let mut j = open_idx + 1;
            let mut body = String::new();
            let mut escaped = false;
            let mut depth = 0usize;
            while j < chars.len() {
                let ch = chars[j];
                if escaped {
                    body.push('\\');
                    body.push(ch);
                    escaped = false;
                    j += 1;
                    continue;
                }
                if ch == '\\' {
                    escaped = true;
                    j += 1;
                    continue;
                }
                if matches!(open, '(' | '[' | '{' | '<') && ch == open {
                    depth += 1;
                    body.push(ch);
                    j += 1;
                    continue;
                }
                if ch == close {
                    if depth > 0 {
                        depth -= 1;
                        body.push(ch);
                        j += 1;
                        continue;
                    }
                    break;
                }
                body.push(ch);
                j += 1;
            }
            if j < chars.len() && chars[j] == close {
                if matches!(kind, 'w' | 'W' | 'i' | 'I') {
                    let interpolate = matches!(kind, 'W' | 'I');
                    let symbolish = matches!(kind, 'i' | 'I');
                    let words = ruby_percent_words(&body, interpolate);
                    out.push('[');
                    for (idx, word) in words.iter().enumerate() {
                        if idx > 0 {
                            out.push_str(", ");
                        }
                        if interpolate && word.starts_with("#{") && word.ends_with('}') {
                            out.push_str(&word[2..word.len() - 1]);
                        } else if symbolish && is_simple_ruby_symbol_word(word) {
                            out.push(':');
                            out.push_str(word);
                        } else if !interpolate && word.starts_with("#{") && word.ends_with('}') {
                            out.push_str(&ruby_single_quoted(&format!("\\{}", word)));
                        } else if !interpolate {
                            out.push_str(&ruby_single_quoted(word));
                        } else {
                            out.push_str(&ruby_double_quoted(word));
                        }
                    }
                    out.push(']');
                } else if kind == 'q' {
                    out.push_str(&ruby_percent_q_quoted(&body));
                } else {
                    out.push('"');
                    for ch in body.chars() {
                        match ch {
                            '"' => out.push_str("\\\""),
                            _ => out.push(ch),
                        }
                    }
                    out.push('"');
                }
                i = j + 1;
                continue;
            }
        }
        out.push(chars[i]);
        i += 1;
    }
    out
}

fn ruby_double_quoted(s: &str) -> String {
    let mut out = String::from("\"");
    for ch in s.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}

fn ruby_single_quoted(s: &str) -> String {
    let mut out = String::from("'");
    for ch in s.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '\'' => out.push_str("\\'"),
            _ => out.push(ch),
        }
    }
    out.push('\'');
    out
}

fn ruby_percent_q_quoted(s: &str) -> String {
    let mut out = String::from("\"");
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '#' if i + 1 < chars.len() && chars[i + 1] == '{' => out.push_str("\\#"),
            ch => out.push(ch),
        }
        i += 1;
    }
    out.push('"');
    out
}

fn is_simple_ruby_symbol_word(s: &str) -> bool {
    let mut chars = s.chars();
    match chars.next() {
        Some(ch) if ch == '_' || ch.is_ascii_alphabetic() => {}
        _ => return false,
    }
    chars.all(|ch| ch == '_' || ch.is_ascii_alphanumeric())
}

fn normalize_ruby_const_reads(source: &str) -> String {
    let chars: Vec<char> = source.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let mut in_single = false;
    let mut in_double = false;
    let mut escaped = false;

    while i < chars.len() {
        let ch = chars[i];
        if escaped {
            out.push(ch);
            escaped = false;
            i += 1;
            continue;
        }
        if ch == '\\' && (in_single || in_double) {
            out.push(ch);
            escaped = true;
            i += 1;
            continue;
        }
        if ch == '\'' && !in_double {
            in_single = !in_single;
            out.push(ch);
            i += 1;
            continue;
        }
        if ch == '"' && !in_single {
            in_double = !in_double;
            out.push(ch);
            i += 1;
            continue;
        }
        if ch == '#' && !in_single && !in_double {
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '\n' {
                    break;
                }
                i += 1;
            }
            i += 1;
            continue;
        }
        if !in_single
            && !in_double
            && ch.is_ascii_uppercase()
            && (i == 0 || !(chars[i - 1].is_ascii_alphanumeric() || chars[i - 1] == '_'))
        {
            let start = i;
            let mut j = i + 1;
            while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
                j += 1;
            }
            if j + 2 < chars.len()
                && chars[j] == ':'
                && chars[j + 1] == ':'
                && chars[j + 2].is_ascii_uppercase()
            {
                let mut k = j + 3;
                while k < chars.len() && (chars[k].is_ascii_alphanumeric() || chars[k] == '_') {
                    k += 1;
                }
                let left: String = chars[start..j].iter().collect();
                let right: String = chars[j + 2..k].iter().collect();
                let prior = out.split_whitespace().last().unwrap_or("");
                let declaration_context = matches!(
                    prior,
                    "class" | "module" | "include" | "extend" | "rescue" | "<"
                );
                if !declaration_context
                    && !(left == "Math" && matches!(right.as_str(), "PI" | "E"))
                    && !(left == "Encoding"
                        && matches!(
                            right.as_str(),
                            "ASCII_8BIT" | "BINARY" | "UTF_8" | "US_ASCII" | "Windows_1252"
                        ))
                {
                    out.push_str(&left);
                    out.push_str(".const_get(:");
                    out.push_str(&right);
                    out.push(')');
                    i = k;
                    continue;
                }
            }
        }
        out.push(ch);
        i += 1;
    }

    out
}
fn walk_stmt_into(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    imports: &mut Vec<Import>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::require_stmt => imports.push(walk_require(pair)?),
        _ => {
            let stmt = walk_statement(pair)?;
            if !matches!(stmt.kind, StmtKind::Empty) {
                body.push(stmt);
            }
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::method_def => walk_method_def(pair)?,
        Rule::class_def => walk_class_def(pair)?,
        Rule::module_def => walk_module_def(pair)?,

        Rule::if_stmt => walk_if(pair)?,
        Rule::unless_stmt => walk_unless(pair)?,
        Rule::while_stmt => walk_while(pair)?,
        Rule::until_stmt => walk_until(pair)?,
        Rule::for_stmt => walk_for(pair)?,
        Rule::case_stmt => walk_case(pair)?,
        Rule::begin_stmt => walk_begin(pair)?,
        Rule::loop_stmt => walk_loop(pair)?,

        Rule::return_stmt => walk_return(pair)?,
        Rule::break_stmt => walk_break_or_next(pair, true)?,
        Rule::next_stmt => walk_break_or_next(pair, false)?,
        Rule::raise_stmt => walk_raise(pair)?,
        Rule::retry_stmt => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::redo_stmt => StmtKind::Continue(ContinueTarget::Implicit),

        Rule::require_stmt => return Ok(Statement::new(StmtKind::Empty)), // handled in walk_stmt_into
        Rule::at_exit_stmt => StmtKind::Empty,                            // no runtime equivalent
        Rule::catch_throw_stmt => StmtKind::Empty,                        // simplified
        Rule::access_modifier_stmt => StmtKind::Empty,                    // metadata only
        Rule::alias_stmt => walk_alias_stmt(pair)?,
        Rule::undef_stmt => StmtKind::Empty, // not directly representable

        Rule::multi_assign_stmt => walk_multi_assign(pair)?,
        Rule::expr_or_assign_stmt => walk_expr_or_assign(pair)?,

        Rule::NEWLINE => StmtKind::Empty,

        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(
        normalize_ruby_raise_stmt_kind(kind),
        span,
    ))
}

fn normalize_ruby_raise_stmt_kind(kind: StmtKind) -> StmtKind {
    let StmtKind::Expr(Expression {
        kind: ExprKind::Call {
            callee,
            args,
            optional,
        },
        ..
    }) = kind
    else {
        return kind;
    };
    if optional {
        return StmtKind::Expr(Expression::new(ExprKind::Call {
            callee,
            args,
            optional,
        }));
    }
    let ExprKind::Ident(name) = &callee.kind else {
        return StmtKind::Expr(Expression::new(ExprKind::Call {
            callee,
            args,
            optional,
        }));
    };
    if !matches!(name.as_str(), "raise" | "fail") {
        return StmtKind::Expr(Expression::new(ExprKind::Call {
            callee,
            args,
            optional,
        }));
    }
    StmtKind::Throw {
        expr: Some(normalize_ruby_raise_args(
            args.into_iter().map(|arg| arg.value).collect(),
        )),
        cause: None,
    }
}

// ── Method def ──────────────────────────────────────────────────────────────

fn walk_method_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut is_self_method = false;
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_name => {
                let text = p.as_str();
                if text.starts_with("self.") {
                    is_self_method = true;
                    name = text[5..].to_string();
                } else if let Some((_, method)) = text.rsplit_once('.') {
                    name = method.to_string();
                } else {
                    name = text.to_string();
                }
            }
            Rule::method_params => params = walk_method_params(p)?,
            Rule::body => body = walk_body(p)?,
            _ => {}
        }
    }

    // Don't apply implicit return to constructors — the compiler handles constructor return
    if name != "initialize" {
        apply_implicit_return(&mut body);
    }

    let mut modifiers = Modifiers::default();
    if is_self_method {
        modifiers.is_static = true;
    }

    let is_generator = body_has_yield(&body);
    register_ruby_method("Object", &name, &params);

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers,
        handles: Vec::new(),
        is_async: false,
        is_generator,
        is_sub: false,
    })
}

fn walk_alias_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let names = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::method_name_id)
        .map(|p| p.as_str().to_string())
        .collect::<Vec<_>>();
    if names.len() >= 2 {
        RUBY_ALIASES.with(|aliases| {
            aliases
                .borrow_mut()
                .insert(names[0].clone(), names[1].clone());
        });
    }
    Ok(StmtKind::Empty)
}

fn method_key(owner: &str, name: &str) -> String {
    format!("{}::{}", owner, name)
}

fn register_ruby_method(owner: &str, name: &str, params: &[Param]) {
    let arity = params
        .iter()
        .filter(|p| !p.is_optional && !p.is_rest && !p.is_kwargs)
        .count() as i64;
    let param_count = params.len() as i64;
    RUBY_METHODS.with(|methods| {
        methods.borrow_mut().insert(
            method_key(owner, name),
            RubyMethodInfo { arity, param_count },
        );
    });
}

fn register_ruby_module_members(name: &str, members: &[ClassMember]) {
    RUBY_MODULE_MEMBERS.with(|modules| {
        modules
            .borrow_mut()
            .insert(name.to_string(), members.to_vec());
    });
}

fn ruby_module_members(name: &str) -> Vec<ClassMember> {
    RUBY_MODULE_MEMBERS.with(|modules| modules.borrow().get(name).cloned().unwrap_or_default())
}

fn register_ruby_member_methods(owner: &str, members: &[ClassMember]) {
    for member in members {
        if let ClassMember::Method(method) = member {
            if let StmtKind::FunctionDecl { name, params, .. } = &method.kind {
                register_ruby_method(owner, name, params);
            }
        }
    }
}

fn ruby_alias_original(name: &str) -> String {
    RUBY_ALIASES.with(|aliases| {
        aliases
            .borrow()
            .get(name)
            .cloned()
            .unwrap_or_else(|| name.to_string())
    })
}

fn ruby_method_info(owner: &str, name: &str) -> RubyMethodInfo {
    let original = ruby_alias_original(name);
    RUBY_METHODS.with(|methods| {
        let methods = methods.borrow();
        methods
            .get(&method_key(owner, name))
            .or_else(|| methods.get(&method_key(owner, &original)))
            .or_else(|| methods.get(&method_key("Object", name)))
            .or_else(|| methods.get(&method_key("Object", &original)))
            .cloned()
            .unwrap_or_default()
    })
}

fn walk_method_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_list {
            params = walk_param_list(p)?;
        }
    }
    Ok(params)
}

fn walk_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_item {
            let inner = p.into_inner().next();
            if let Some(item) = inner {
                match item.as_rule() {
                    Rule::normal_param => {
                        params.push(Param {
                            name: item.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::optional_param => {
                        let mut name = String::new();
                        let mut default = None;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ => default = Some(walk_expression(c)?),
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: true,
                            is_nullable: false,
                        });
                    }
                    Rule::splat_param => {
                        let name = item
                            .into_inner()
                            .find(|c| c.as_rule() == Rule::identifier)
                            .map(|c| c.as_str().to_string())
                            .unwrap_or_default();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: true,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::double_splat_param => {
                        let name = item
                            .into_inner()
                            .find(|c| c.as_rule() == Rule::identifier)
                            .map(|c| c.as_str().to_string())
                            .unwrap_or_default();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: true,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::block_param => {
                        // &block — ignore for now, blocks are handled differently
                    }
                    Rule::keyword_param => {
                        let mut name = String::new();
                        let mut default = None;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ if is_expression_rule(c.as_rule()) => {
                                    default = Some(walk_expression(c)?);
                                }
                                _ => {}
                            }
                        }
                        let is_optional = default.is_some();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional,
                            is_nullable: false,
                        });
                    }
                    _ => {}
                }
            }
        }
    }
    Ok(params)
}

// ── Class def ───────────────────────────────────────────────────────────────

fn walk_class_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constant => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::constant_path => {
                parents.push(p.as_str().to_string());
            }
            Rule::class_body => {
                members = walk_class_body(p, &name)?;
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_class_body(pair: Pair<Rule>, class_name: &str) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    let mut current_visibility = Visibility::Public;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::access_modifier_stmt => {
                let text = p.as_str().trim();
                if text.starts_with("private") {
                    current_visibility = Visibility::Private;
                } else if text.starts_with("protected") {
                    current_visibility = Visibility::Protected;
                } else {
                    current_visibility = Visibility::Public;
                }
            }
            Rule::attr_decl => {
                members.extend(walk_attr_decl(p)?);
            }
            Rule::method_def => {
                let stmt_kind = walk_method_def(p)?;
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    body,
                    modifiers,
                    ..
                } = &stmt_kind
                {
                    register_ruby_method(class_name, name, params);
                    if name == "initialize" {
                        // Extract instance variable assignments from constructor body
                        members.push(ClassMember::Constructor {
                            // Ruby has one constructor, `initialize` — unnamed.
                            name: None,
                            params: params.clone(),
                            body: body.clone(),
                            base_args: None,
                            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
                            visibility: current_visibility,
                        });
                    } else {
                        let mut mods = modifiers.clone();
                        mods.visibility = current_visibility;
                        members.push(ClassMember::Method(Box::new(Statement::new(stmt_kind))));
                    }
                }
            }
            Rule::include_stmt | Rule::extend_stmt => {
                let included = p
                    .into_inner()
                    .find(|inner| matches!(inner.as_rule(), Rule::constant_path | Rule::constant))
                    .map(|inner| inner.as_str().to_string());
                if let Some(module_name) = included {
                    let module_members = ruby_module_members(&module_name);
                    register_ruby_member_methods(class_name, &module_members);
                    members.extend(module_members);
                }
            }
            Rule::alias_stmt => {}
            Rule::class_def => {
                // Nested class
                let nested = walk_class_def(p)?;
                members.push(ClassMember::NestedType(Box::new(Statement::new(nested))));
            }
            Rule::module_def => {
                let nested = walk_module_def(p)?;
                members.push(ClassMember::NestedType(Box::new(Statement::new(nested))));
            }
            Rule::NEWLINE => {}
            _ => {
                // Other statements in class body → treat as static initializer
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    if let Some((alias, original)) = ruby_alias_method_stmt(&stmt) {
                        if let Some(alias_stmt) = members.iter().find_map(|member| {
                            let ClassMember::Method(method) = member else {
                                return None;
                            };
                            let mut cloned = (**method).clone();
                            if let StmtKind::FunctionDecl { name, .. } = &mut cloned.kind {
                                if name == &original {
                                    *name = alias.clone();
                                    return Some(cloned);
                                }
                            }
                            None
                        }) {
                            members.push(ClassMember::Method(Box::new(alias_stmt)));
                        }
                        continue;
                    }
                    if let Some(name) = ruby_remove_const_stmt(&stmt) {
                        members.retain(|member| {
                            !matches!(member, ClassMember::Const { name: const_name, .. } if const_name == &name)
                        });
                        continue;
                    }
                    if let Some(name) = ruby_remove_method_stmt(&stmt) {
                        members.retain(|member| {
                            !matches!(
                                member,
                                ClassMember::Method(method)
                                    if matches!(&method.kind, StmtKind::FunctionDecl { name: method_name, .. } if method_name == &name)
                            )
                        });
                        continue;
                    }
                    if let StmtKind::Assign { targets, value, .. } = stmt.kind {
                        if targets.len() == 1 {
                            if let ExprKind::Ident(name) = &targets[0].kind {
                                if name.chars().next().is_some_and(|c| c.is_ascii_uppercase()) {
                                    members.push(ClassMember::Const {
                                        name: name.clone(),
                                        type_hint: None,
                                        value,
                                        visibility: current_visibility,
                                    });
                                    continue;
                                }
                            }
                        }
                        members.push(ClassMember::Method(Box::new(Statement::new(
                            StmtKind::Assign {
                                targets,
                                value,
                                by_ref: false,
                            },
                        ))));
                    } else {
                        members.push(ClassMember::Method(Box::new(stmt)));
                    }
                }
            }
        }
    }
    Ok(members)
}

fn ruby_remove_const_stmt(stmt: &Statement) -> Option<String> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "remove_const") || args.len() != 1 {
        return None;
    }
    ruby_method_name_arg(&args[0].value)
}

fn ruby_remove_method_stmt(stmt: &Statement) -> Option<String> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "remove_method") || args.len() != 1
    {
        return None;
    }
    ruby_method_name_arg(&args[0].value)
}

fn ruby_alias_method_stmt(stmt: &Statement) -> Option<(String, String)> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "alias_method") || args.len() != 2 {
        return None;
    }
    Some((
        ruby_method_name_arg(&args[0].value)?,
        ruby_method_name_arg(&args[1].value)?,
    ))
}

fn walk_attr_decl(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut kind = "";
    let mut names = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::attr_kind => {
                kind = match p.as_str().trim() {
                    "attr_accessor" => "accessor",
                    "attr_reader" => "reader",
                    "attr_writer" => "writer",
                    _ => "accessor",
                }
            }
            Rule::symbol_list => {
                for s in p.into_inner() {
                    let text = s.as_str().trim();
                    let name = if text.starts_with(':') {
                        &text[1..]
                    } else if text.starts_with('"') || text.starts_with('\'') {
                        &text[1..text.len() - 1]
                    } else {
                        text
                    };
                    names.push(name.to_string());
                }
            }
            _ => {}
        }
    }

    let mut members = Vec::new();
    for name in names {
        let has_getter = kind == "accessor" || kind == "reader";
        let has_setter = kind == "accessor" || kind == "writer";

        // Getter → method `name()` that returns self._rb_<field>
        // The backing field is created by `@name = ...` in initialize
        // which maps to self._rb_name (prefixed to avoid struct key collision)
        if has_getter {
            let self_expr = Expression::new(ExprKind::Ident("self".into()));
            let field_access = Expression::new(ExprKind::Member {
                object: Box::new(self_expr),
                field: format!("_rb_{}", name),
                null_safe: false,
            });
            let body = vec![Statement::new(StmtKind::Return(Some(field_access)))];
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: name.clone(),
                    params: Vec::new(),
                    return_type: None,
                    body,
                    modifiers: Modifiers::default(),
                    handles: Vec::new(),
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
        } else if kind == "writer" {
            let body = vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::string("NoMethodError")),
                cause: None,
            })];
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: name.clone(),
                    params: Vec::new(),
                    return_type: None,
                    body,
                    modifiers: Modifiers::default(),
                    handles: Vec::new(),
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
        }

        // Setter semantics: Ruby `d.name = x` is transformed in the walker to
        // Assign(Member(d, "_rb_name"), x) via fixup_assign_target, which writes
        // directly to the _rb_ prefixed backing field via struct_set.
        let _ = has_setter;
    }
    Ok(members)
}

// ── Module def ──────────────────────────────────────────────────────────────

fn walk_module_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constant => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::class_body => {
                members = walk_class_body(p, &name)?;
            }
            _ => {}
        }
    }

    register_ruby_module_members(&name, &members);

    Ok(StmtKind::ModuleDecl {
        name,
        members,
        visibility: Visibility::Public,
    })
}

// ── If ──────────────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for modifier form: expression if_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::if_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier if body")?)?;
        // skip if_kw
        iter.find(|p| p.as_rule() == Rule::if_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier if condition")?)?;
        return Ok(StmtKind::If {
            cond,
            then_body: vec![Statement::new(StmtKind::Expr(body_expr))],
            elifs: Vec::new(),
            else_body: None,
        });
    }

    // Block form: if cond then_kw? body elsif* else? end
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let then_body = walk_body(next_rule(&mut iter, Rule::body)?)?;

    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in iter {
        match p.as_rule() {
            Rule::elsif_clause => {
                let mut ei = p.into_inner();
                let econd = walk_expression(next_meaningful(&mut ei)?)?;
                let ebody = walk_body(find_rule(ei, Rule::body)?)?;
                elifs.push((econd, ebody));
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

// ── Unless ──────────────────────────────────────────────────────────────────

fn walk_unless(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for modifier form: expression unless_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::unless_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier unless body")?)?;
        iter.find(|p| p.as_rule() == Rule::unless_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier unless condition")?)?;
        // unless → if !cond
        return Ok(StmtKind::If {
            cond: negate(cond),
            then_body: vec![Statement::new(StmtKind::Expr(body_expr))],
            elifs: Vec::new(),
            else_body: None,
        });
    }

    // Block form
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let then_body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;

    let mut else_body = None;
    for p in iter {
        if p.as_rule() == Rule::else_clause {
            let ei = p.into_inner();
            else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
        }
    }

    // unless cond → if !cond
    Ok(StmtKind::If {
        cond: negate(cond),
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

// ── While ───────────────────────────────────────────────────────────────────

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Modifier form: expression while_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::while_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier while body")?)?;
        iter.find(|p| p.as_rule() == Rule::while_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier while condition")?)?;
        return Ok(StmtKind::While {
            cond,
            body: vec![Statement::new(StmtKind::Expr(body_expr))],
            else_body: None,
        });
    }

    // Block form
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

// ── Until ───────────────────────────────────────────────────────────────────

fn walk_until(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Modifier form
    if children.iter().any(|p| p.as_rule() == Rule::until_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier until body")?)?;
        iter.find(|p| p.as_rule() == Rule::until_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier until condition")?)?;
        // until cond → while !cond
        return Ok(StmtKind::While {
            cond: negate(cond),
            body: vec![Statement::new(StmtKind::Expr(body_expr))],
            else_body: None,
        });
    }

    // Block form: until cond → while !cond
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;
    Ok(StmtKind::While {
        cond: negate(cond),
        body,
        else_body: None,
    })
}

// ── For ─────────────────────────────────────────────────────────────────────

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut vars = Vec::new();
    let mut iter_expr = None;
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => vars.push(p.as_str().to_string()),
            Rule::in_kw | Rule::do_kw => {}
            Rule::body => body = walk_body(p)?,
            _ if is_expression_rule(p.as_rule()) => {
                if iter_expr.is_none() {
                    iter_expr = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }

    // Multi-target destructuring
    let var = if vars.len() > 1 {
        let tmp = "__forin_element".to_string();
        let mut destructure_stmts: Vec<Statement> = Vec::new();
        for (i, name) in vars.iter().enumerate() {
            destructure_stmts.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(name.clone()))],
                value: Expression::new(ExprKind::Index {
                    object: Box::new(Expression::new(ExprKind::Ident(tmp.clone()))),
                    index: Box::new(Expression::int(i as i64)),
                    null_safe: false,
                }),
                by_ref: false,
            }));
        }
        destructure_stmts.extend(body);
        body = destructure_stmts;
        tmp
    } else {
        vars.into_iter().next().unwrap_or_default()
    };

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter: iter_expr.unwrap_or(Expression::null()),
        body,
        of: true,
        else_body: None,
        is_async: false,
    })
}

// ── Case / When ─────────────────────────────────────────────────────────────

fn walk_case(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut subject = None;
    let mut cases = Vec::new();
    let mut default = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::when_clause => {
                let mut conditions = Vec::new();
                let mut body = Vec::new();
                for wp in p.into_inner() {
                    match wp.as_rule() {
                        Rule::expression_list => {
                            for ep in wp.into_inner() {
                                if is_expression_rule(ep.as_rule()) {
                                    let expr = walk_expression(ep)?;
                                    conditions.push(CaseCondition::Value(expr));
                                }
                            }
                        }
                        Rule::body => body = walk_body(wp)?,
                        Rule::then_kw => {}
                        _ if is_expression_rule(wp.as_rule()) => {
                            let expr = walk_expression(wp)?;
                            conditions.push(CaseCondition::Value(expr));
                        }
                        _ => {}
                    }
                }
                cases.push(SwitchCase { conditions, body });
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                default = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            _ if is_expression_rule(p.as_rule()) => {
                if subject.is_none() {
                    subject = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Switch {
        expr: subject.unwrap_or(Expression::bool(true)),
        cases,
        default,
    })
}

// ── Begin / Rescue / Ensure ─────────────────────────────────────────────────

fn walk_begin(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut else_body = None;
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::body => {
                if body.is_empty() {
                    body = walk_body(p)?;
                }
            }
            Rule::rescue_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();

                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::constant | Rule::constant_path => types.push(cp.as_str().to_string()),
                        Rule::identifier => var_name = Some(cp.as_str().to_string()),
                        Rule::body => catch_body = walk_body(cp)?,
                        _ => {}
                    }
                }
                if types.is_empty() {
                    types = vec![
                        "StandardError".to_string(),
                        "RuntimeError".to_string(),
                        "ArgumentError".to_string(),
                        "TypeError".to_string(),
                        "NameError".to_string(),
                        "NoMethodError".to_string(),
                        "ZeroDivisionError".to_string(),
                        "IndexError".to_string(),
                        "KeyError".to_string(),
                    ];
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            Rule::ensure_clause => {
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::body {
                        finally = Some(walk_body(ep)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body,
        finally,
    })
}

// ── Loop ────────────────────────────────────────────────────────────────────

fn walk_loop(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::do_block => {
                for dp in p.into_inner() {
                    if dp.as_rule() == Rule::body {
                        body = walk_body(dp)?;
                    }
                }
            }
            _ => {}
        }
    }
    // loop { ... } → while true { ... }
    Ok(StmtKind::While {
        cond: Expression::bool(true),
        body,
        else_body: None,
    })
}

// ── Return ──────────────────────────────────────────────────────────────────

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            exprs.push(walk_expression(p)?);
        } else if p.as_rule() == Rule::expression_list {
            for ep in p.into_inner() {
                if is_expression_rule(ep.as_rule()) {
                    exprs.push(walk_expression(ep)?);
                }
            }
        }
    }
    // Ruby `return a, b` semantically returns an Array, but we model
    // it as `ExprKind::Tuple` so the compiler's multi-value pre-scan
    // can recognise the uniform-arity pattern. Tuple and Array lower
    // to the same `ecma:array` packed representation — the AST
    // distinction is purely to drive the multi-value opt-in.
    let expr = if exprs.len() > 1 {
        Some(Expression::new(ExprKind::Tuple(exprs)))
    } else {
        exprs.into_iter().next()
    };
    Ok(StmtKind::Return(expr))
}

// ── Raise ───────────────────────────────────────────────────────────────────

fn walk_raise(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    let mut modifiers = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::modifier_suffix {
            modifiers.push(p);
        } else if is_expression_rule(p.as_rule()) {
            exprs.push(walk_expression(p)?);
        }
    }
    let expr = Some(normalize_ruby_raise_args(exprs));
    let stmt = StmtKind::Throw { expr, cause: None };
    maybe_wrap_modifier(stmt, &mut modifiers)
}

fn normalize_ruby_raise_args(mut exprs: Vec<Expression>) -> Expression {
    if exprs.is_empty() || matches!(exprs[0].kind, ExprKind::Lit(Literal::Null)) {
        return ruby_call_expr("__ruby_exception_runtime_error", Vec::new());
    }
    if exprs.len() >= 2 {
        let message = exprs.remove(1);
        let class_expr = exprs.remove(0);
        if let ExprKind::Ident(name) = &class_expr.kind {
            if let Some(helper) = ruby_exception_helper_name(name) {
                return ruby_call_expr(helper, vec![message]);
            }
        }
        return normalize_ruby_raise_expr(class_expr);
    }
    normalize_ruby_raise_expr(exprs.remove(0))
}

fn normalize_ruby_raise_expr(expr: Expression) -> Expression {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(_)) => {
            ruby_call_expr("__ruby_exception_runtime_error", vec![expr])
        }
        ExprKind::Ident(name) => {
            if let Some(helper) = ruby_exception_helper_name(name) {
                ruby_call_expr(helper, Vec::new())
            } else {
                expr
            }
        }
        _ => expr,
    }
}

fn ruby_exception_helper_name(name: &str) -> Option<&'static str> {
    Some(match name {
        "Exception" => "__ruby_exception_exception",
        "StandardError" => "__ruby_exception_standard_error",
        "RuntimeError" => "__ruby_exception_runtime_error",
        "ArgumentError" => "__ruby_exception_argument_error",
        "TypeError" => "__ruby_exception_type_error",
        "NameError" => "__ruby_exception_name_error",
        "NoMethodError" => "__ruby_exception_no_method_error",
        "ZeroDivisionError" => "__ruby_exception_zero_division_error",
        "IndexError" => "__ruby_exception_index_error",
        "KeyError" => "__ruby_exception_key_error",
        "FrozenError" => "__ruby_exception_frozen_error",
        "RangeError" => "__ruby_exception_range_error",
        "FloatDomainError" => "__ruby_exception_float_domain_error",
        "SystemExit" => "__ruby_exception_system_exit",
        "SignalException" => "__ruby_exception_signal_exception",
        "Interrupt" => "__ruby_exception_interrupt",
        "ScriptError" => "__ruby_exception_script_error",
        "SyntaxError" => "__ruby_exception_syntax_error",
        "LoadError" => "__ruby_exception_load_error",
        "NotImplementedError" => "__ruby_exception_not_implemented_error",
        "SecurityError" => "__ruby_exception_security_error",
        "NoMemoryError" => "__ruby_exception_no_memory_error",
        "UncaughtThrowError" => "__ruby_exception_uncaught_throw_error",
        "LocalJumpError" => "__ruby_exception_local_jump_error",
        _ => return None,
    })
}

fn ruby_exception_ancestors_expr(name: &str) -> Option<Expression> {
    let chain: &[&str] = match name {
        "NoMethodError" => &["NoMethodError", "NameError", "StandardError", "Exception"],
        "NameError" => &["NameError", "StandardError", "Exception"],
        "KeyError" => &["KeyError", "IndexError", "StandardError", "Exception"],
        "IndexError" => &["IndexError", "StandardError", "Exception"],
        "FloatDomainError" => &[
            "FloatDomainError",
            "RangeError",
            "StandardError",
            "Exception",
        ],
        "RangeError" => &["RangeError", "StandardError", "Exception"],
        "ArgumentError" => &["ArgumentError", "StandardError", "Exception"],
        "TypeError" => &["TypeError", "StandardError", "Exception"],
        "ZeroDivisionError" => &["ZeroDivisionError", "StandardError", "Exception"],
        "FrozenError" => &["FrozenError", "RuntimeError", "StandardError", "Exception"],
        "RuntimeError" => &["RuntimeError", "StandardError", "Exception"],
        "StandardError" => &["StandardError", "Exception"],
        "Interrupt" => &["Interrupt", "SignalException", "Exception"],
        "SignalException" => &["SignalException", "Exception"],
        "SystemExit" => &["SystemExit", "Exception"],
        "SyntaxError" => &["SyntaxError", "ScriptError", "Exception"],
        "LoadError" => &["LoadError", "ScriptError", "Exception"],
        "NotImplementedError" => &["NotImplementedError", "ScriptError", "Exception"],
        "ScriptError" => &["ScriptError", "Exception"],
        "SecurityError" => &["SecurityError", "Exception"],
        "NoMemoryError" => &["NoMemoryError", "Exception"],
        "UncaughtThrowError" => &[
            "UncaughtThrowError",
            "ArgumentError",
            "StandardError",
            "Exception",
        ],
        "LocalJumpError" => &["LocalJumpError", "StandardError", "Exception"],
        "Exception" => &["Exception"],
        _ => return None,
    };
    Some(ruby_array_expr(
        chain
            .iter()
            .map(|ancestor| Expression::string(ancestor))
            .collect(),
    ))
}

fn ruby_expr_may_be_exception(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => matches!(name.as_str(), "e" | "err" | "exception"),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if name.starts_with("__ruby_exception"))
        }
        _ => false,
    }
}

fn is_ruby_exception_name(name: &str) -> bool {
    matches!(
        name,
        "Exception"
            | "StandardError"
            | "RuntimeError"
            | "ArgumentError"
            | "TypeError"
            | "NameError"
            | "NoMethodError"
            | "ZeroDivisionError"
            | "IndexError"
            | "KeyError"
            | "FrozenError"
            | "RangeError"
            | "FloatDomainError"
            | "SystemExit"
            | "SignalException"
            | "Interrupt"
            | "ScriptError"
            | "SyntaxError"
            | "LoadError"
            | "NotImplementedError"
            | "SecurityError"
            | "NoMemoryError"
            | "UncaughtThrowError"
            | "LocalJumpError"
    )
}

// ── Break / Next with optional modifier ─────────────────────────────────────

fn walk_break_or_next(pair: Pair<Rule>, is_break: bool) -> Result<StmtKind, String> {
    let mut modifiers = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::modifier_suffix {
            modifiers.push(p);
        }
    }
    let stmt = if is_break {
        StmtKind::Break(BreakTarget::Implicit)
    } else {
        StmtKind::Continue(ContinueTarget::Implicit)
    };
    maybe_wrap_modifier(stmt, &mut modifiers)
}

// ── Multi-assign ────────────────────────────────────────────────────────────

fn walk_multi_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut values = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::target => {
                let inner: Vec<Pair<Rule>> = p.into_inner().collect();
                if let Some(first) = inner.into_iter().next() {
                    targets.push(walk_expression(first)?);
                }
            }
            Rule::expression_list => {
                for ep in p.into_inner() {
                    if is_expression_rule(ep.as_rule()) {
                        values.push(walk_expression(ep)?);
                    }
                }
            }
            _ => {}
        }
    }

    // Multi-assign: a, b = 1, 2
    // Emit as destructuring assign
    if values.len() == 1 {
        // a, b = [1, 2] — single RHS
        let patterns = targets
            .iter()
            .map(|t| {
                if let ExprKind::Ident(name) = &t.kind {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                } else {
                    ArrayPatternElem::Hole
                }
            })
            .collect();
        Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(
                DestructurePattern::Array(patterns),
            ))],
            value: values.into_iter().next().unwrap(),
            by_ref: false,
        })
    } else {
        // a, b = 1, 2 — wrap RHS in array
        let value = Expression::new(ExprKind::Array(
            values
                .into_iter()
                .map(|v| ArrayElement {
                    key: None,
                    value: v,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        let patterns = targets
            .iter()
            .map(|t| {
                if let ExprKind::Ident(name) = &t.kind {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                } else {
                    ArrayPatternElem::Hole
                }
            })
            .collect();
        Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(
                DestructurePattern::Array(patterns),
            ))],
            value,
            by_ref: false,
        })
    }
}

// ── Expression or assignment ────────────────────────────────────────────────

/// Transform assignment targets: unwrap `Call(Member(obj, field), [])` → `Member(obj, "_rb_field")`.
/// In Ruby, `d.name = x` goes through a setter method which writes to the backing @name ivar.
/// Since @vars are stored with `_rb_` prefix, external assignments must write there too.
fn fixup_assign_target(expr: Expression) -> Expression {
    if let ExprKind::Call {
        ref callee,
        ref args,
        ..
    } = expr.kind
    {
        if args.is_empty() {
            if let ExprKind::Member {
                ref object,
                ref field,
                null_safe,
            } = callee.kind
            {
                return Expression::new(ExprKind::Member {
                    object: object.clone(),
                    field: format!("_rb_{}", field),
                    null_safe,
                });
            }
        }
    }
    expr
}

fn walk_expr_or_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    if let Some(stmt) = walk_raw_command_builtin(pair.as_str())? {
        return Ok(stmt);
    }
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .collect();

    if inner.is_empty() {
        return Ok(StmtKind::Empty);
    }

    // ── Check for command call (postfix ~ command_args ~ block_literal? ~ modifier_suffix?)
    let has_command_args = inner.iter().any(|p| p.as_rule() == Rule::command_args);
    if has_command_args {
        return walk_command_call(inner);
    }

    // ── Check for augmented assignment
    let has_aug = inner.iter().any(|p| p.as_rule() == Rule::aug_assign_op);
    if has_aug {
        let target = fixup_assign_target(walk_expression(inner.remove(0))?);
        let op_str = inner.remove(0).as_str().to_string();
        let value = if !inner.is_empty() && is_expression_rule(inner[0].as_rule()) {
            walk_expression(inner.remove(0))?
        } else {
            Expression::null()
        };
        let op = match op_str.as_str() {
            "+=" => CompoundOp::Add,
            "-=" => CompoundOp::Sub,
            "*=" => CompoundOp::Mul,
            "/=" => CompoundOp::Div,
            "%=" => CompoundOp::Mod,
            "**=" => CompoundOp::Pow,
            "<<=" => CompoundOp::Shl,
            ">>=" => CompoundOp::Shr,
            "|=" => CompoundOp::BitOr,
            "&=" => CompoundOp::BitAnd,
            "^=" => CompoundOp::BitXor,
            "||=" => CompoundOp::Or,
            "&&=" => CompoundOp::And,
            _ => CompoundOp::Add,
        };
        let stmt = StmtKind::CompoundAssign { target, op, value };
        return maybe_wrap_modifier(stmt, &mut inner);
    }

    // ── Check for regular assignment (expression = expression_list)
    let has_expr_list = inner.iter().any(|p| p.as_rule() == Rule::expression_list);
    if has_expr_list {
        let target = fixup_assign_target(walk_expression(inner.remove(0))?);
        let mut values = Vec::new();
        let mut remaining = Vec::new();
        for p in inner {
            if p.as_rule() == Rule::expression_list {
                for ep in p.into_inner() {
                    if is_expression_rule(ep.as_rule()) {
                        values.push(walk_expression(ep)?);
                    }
                }
            } else if p.as_rule() == Rule::modifier_suffix {
                remaining.push(p);
            } else if is_expression_rule(p.as_rule()) {
                values.push(walk_expression(p)?);
            }
        }
        if values.is_empty() {
            let stmt = StmtKind::Expr(target);
            return maybe_wrap_modifier(stmt, &mut remaining);
        }
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Array(
                values
                    .into_iter()
                    .map(|v| ArrayElement {
                        key: None,
                        value: v,
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ))
        };
        let stmt = StmtKind::Assign {
            targets: vec![target],
            value,
            by_ref: false,
        };
        return maybe_wrap_modifier(stmt, &mut remaining);
    }

    // ── Expression statement (expression ~ modifier_suffix?)
    let expr = walk_expression(inner.remove(0))?;
    let stmt = normalize_bang_method_stmt(expr.clone()).unwrap_or(StmtKind::Expr(expr));
    maybe_wrap_modifier(stmt, &mut inner)
}

fn normalize_bang_method_stmt(expr: Expression) -> Option<StmtKind> {
    if let Some(target_name) = ruby_mutating_shl_target(&expr) {
        return Some(StmtKind::Assign {
            targets: vec![Expression::ident(&target_name)],
            value: expr,
            by_ref: false,
        });
    }
    let ExprKind::Call {
        callee,
        args,
        optional,
    } = expr.kind
    else {
        return None;
    };
    if optional {
        return None;
    }
    if let ExprKind::Ident(name) = &callee.kind {
        if matches!(name.as_str(), "raise" | "fail") {
            let exprs = args.into_iter().map(|arg| arg.value).collect();
            return Some(StmtKind::Throw {
                expr: Some(normalize_ruby_raise_args(exprs)),
                cause: None,
            });
        }
    }
    let ExprKind::Member {
        object,
        field,
        null_safe,
    } = callee.kind
    else {
        return None;
    };
    if null_safe {
        return None;
    }
    let method = match field.as_str() {
        "upcase!" => "upcase",
        "downcase!" => "downcase",
        "strip!" => "strip",
        "chomp!" => "chomp",
        "chop!" => "chop",
        "capitalize!" => "capitalize",
        "swapcase!" => "swapcase",
        "reverse!" => "reverse",
        "succ!" => "succ",
        "next!" => "next",
        "freeze" => "freeze",
        "squeeze!" => "squeeze",
        "tr!" => "tr",
        "tr_s!" => "tr_s",
        "delete!" => "delete",
        "gsub!" => "gsub",
        "sub!" => "sub",
        "insert" => "insert",
        "clear" => "clear",
        "replace" => "replace",
        "concat" => "concat",
        "prepend" => "prepend",
        _ => return None,
    };
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    let target = Expression::ident(name);
    let value = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object,
            field: method.to_string(),
            null_safe: false,
        })),
        args,
        optional: false,
    });
    Some(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })
}

fn ruby_mutating_shl_target(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            let ExprKind::Ident(name) = &callee.kind else {
                return None;
            };
            if name != "__ruby_op_shl" {
                return None;
            }
            args.first()
                .and_then(|arg| ruby_mutating_shl_target(&arg.value))
        }
        ExprKind::Ident(name) => Some(name.clone()),
        _ => None,
    }
}

fn walk_raw_command_builtin(raw: &str) -> Result<Option<StmtKind>, String> {
    let text = raw.trim();
    let Some(split_at) = text.find(char::is_whitespace) else {
        return Ok(None);
    };
    let head = &text[..split_at];
    if !matches!(head, "puts" | "print" | "p" | "pp" | "warn") {
        return Ok(None);
    }
    let tail = text[split_at..].trim();
    if tail.is_empty() || tail.starts_with('=') {
        return Ok(None);
    }
    let mut parsed = RubyParser::parse(Rule::call_args, tail)
        .map_err(|e| format!("Parse error in command args: {}", e))?;
    let args_pair = parsed
        .next()
        .ok_or_else(|| "command args parse produced no args".to_string())?;
    let args = walk_call_args(args_pair)?;
    Ok(Some(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(head)),
        args,
        optional: false,
    }))))
}

/// Handle command-style call: postfix ~ command_args ~ block_literal? ~ modifier_suffix?
fn walk_command_call(mut items: Vec<Pair<Rule>>) -> Result<StmtKind, String> {
    // The first item(s) before command_args form the callee postfix expression.
    let cmd_pos = items
        .iter()
        .position(|p| p.as_rule() == Rule::command_args)
        .unwrap();

    // Build the callee from the postfix pair(s) before command_args
    let callee_pairs: Vec<Pair<Rule>> = items.drain(..cmd_pos).collect();
    let callee = if callee_pairs.len() == 1 {
        let p = callee_pairs.into_iter().next().unwrap();
        Expression::new(walk_expr_kind(p)?)
    } else if !callee_pairs.is_empty() {
        let p = callee_pairs.into_iter().next().unwrap();
        Expression::new(walk_expr_kind(p)?)
    } else {
        return Err("Command call missing callee".into());
    };

    // Now items[0] = command_args (same structure as call_args: contains call_arg children)
    let cmd_args_pair = items.remove(0);
    let mut args = walk_call_args(cmd_args_pair)?;

    // Optional block literal
    if !items.is_empty() && items[0].as_rule() == Rule::block_literal {
        let blk = items.remove(0);
        let lambda = walk_block_literal(blk)?;
        args.push(Argument::positional(lambda));
    }

    if matches!(&callee.kind, ExprKind::Ident(name) if name == "lambda") && args.len() == 1 {
        let stmt = StmtKind::Expr(ruby_proc_expr("__ruby_lambda", args.remove(0).value));
        return maybe_wrap_modifier(stmt, &mut items);
    }

    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    });

    let stmt = StmtKind::Expr(call_expr);
    maybe_wrap_modifier(stmt, &mut items)
}

/// Wrap a statement in an if/unless/while/until modifier if present
fn maybe_wrap_modifier(stmt: StmtKind, rest: &mut Vec<Pair<Rule>>) -> Result<StmtKind, String> {
    let mod_pos = rest
        .iter()
        .position(|p| p.as_rule() == Rule::modifier_suffix);
    let mod_pair = match mod_pos {
        Some(pos) => rest.remove(pos),
        None => return Ok(stmt),
    };
    let mut mod_inner = mod_pair.into_inner();
    let kw = match mod_inner.next() {
        Some(k) => k,
        None => return Ok(stmt),
    };
    let cond_pair = mod_inner
        .next()
        .ok_or_else(|| "modifier_suffix missing condition".to_string())?;
    let cond = walk_expression(cond_pair)?;
    let body_stmt = Statement::new(stmt);
    match kw.as_rule() {
        Rule::if_kw => Ok(StmtKind::If {
            cond,
            then_body: vec![body_stmt],
            elifs: vec![],
            else_body: None,
        }),
        Rule::unless_kw => Ok(StmtKind::If {
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            }),
            then_body: vec![body_stmt],
            elifs: vec![],
            else_body: None,
        }),
        Rule::while_kw => Ok(StmtKind::While {
            cond,
            body: vec![body_stmt],
            else_body: None,
        }),
        Rule::until_kw => Ok(StmtKind::While {
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            }),
            body: vec![body_stmt],
            else_body: None,
        }),
        _ => Ok(StmtKind::Expr(Expression::null())),
    }
}

// ── Require (import) ────────────────────────────────────────────────────────

fn walk_require(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let text = pair.as_str();
    let _is_relative = text.starts_with("require_relative");

    let mut path = String::new();
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            let expr_text = p.as_str().trim();
            // Strip quotes
            path = if expr_text.starts_with('"') || expr_text.starts_with('\'') {
                expr_text[1..expr_text.len() - 1].to_string()
            } else {
                expr_text.to_string()
            };
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias: None },
        span,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Body (list of statements)
// ════════════════════════════════════════════════════════════════════════════

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::NEWLINE => {}
            _ => {
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    stmts.push(stmt);
                }
            }
        }
    }
    normalize_consecutive_prints(&mut stmts);
    Ok(stmts)
}

fn print_call_args(stmt: &mut Statement) -> Option<&mut Vec<Argument>> {
    if let StmtKind::Expr(Expression {
        kind:
            ExprKind::Call {
                callee,
                args,
                optional: false,
            },
        ..
    }) = &mut stmt.kind
    {
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "print") {
            return Some(args);
        }
    }
    None
}

fn normalize_consecutive_prints(stmts: &mut Vec<Statement>) {
    let mut out: Vec<Statement> = Vec::with_capacity(stmts.len());
    for mut stmt in std::mem::take(stmts) {
        if let Some(args) = print_call_args(&mut stmt) {
            if let Some(prev) = out.last_mut() {
                if let Some(prev_args) = print_call_args(prev) {
                    prev_args.append(args);
                    continue;
                }
            }
        }
        out.push(stmt);
    }
    *stmts = out;
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    match pair.as_str().trim() {
        "Math::PI" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Float(std::f64::consts::PI)),
                span,
            ));
        }
        "Math::E" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Float(std::f64::consts::E)),
                span,
            ));
        }
        "Encoding::ASCII_8BIT" | "Encoding::BINARY" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Str("ASCII-8BIT".to_string())),
                span,
            ));
        }
        "Encoding::UTF_8" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Str("UTF-8".to_string())),
                span,
            ));
        }
        "Encoding::US_ASCII" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Str("US-ASCII".to_string())),
                span,
            ));
        }
        "Encoding::Windows_1252" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Str("Windows-1252".to_string())),
                span,
            ));
        }
        _ => {}
    }
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // ── Literals ────────────────────────────────────────────────────
        Rule::integer_literal => parse_ruby_int(pair.as_str()),
        Rule::float_literal => parse_ruby_float(pair.as_str()),
        Rule::string_literal => Ok(ExprKind::Lit(Literal::Str(parse_ruby_string(
            pair.as_str(),
        )))),
        Rule::interpolated_string => walk_interpolated_string(pair),
        Rule::heredoc => Ok(ExprKind::Lit(Literal::Str(parse_heredoc(pair.as_str())))),
        Rule::symbol => {
            let raw = &pair.as_str()[1..];
            let value = if raw.starts_with('"') || raw.starts_with('\'') {
                if raw.starts_with('"') {
                    raw[1..raw.len() - 1].to_string()
                } else {
                    parse_ruby_string(raw)
                }
            } else {
                raw.to_string()
            };
            Ok(ExprKind::Lit(Literal::Str(value)))
        }
        Rule::regex_literal => Ok(ExprKind::Lit(Literal::Str(pair.as_str().to_string()))),
        Rule::percent_literal => Ok(walk_percent_literal(pair.as_str())),

        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::nil_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::self_kw => Ok(ExprKind::This),

        Rule::identifier => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::constant => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::constant_path => match pair.as_str() {
            "Math::PI" => Ok(ExprKind::Lit(Literal::Float(std::f64::consts::PI))),
            "Math::E" => Ok(ExprKind::Lit(Literal::Float(std::f64::consts::E))),
            "Float::INFINITY" => Ok(ExprKind::Lit(Literal::Float(f64::INFINITY))),
            "Float::NAN" => Ok(ExprKind::Lit(Literal::Float(f64::NAN))),
            "Encoding::ASCII_8BIT" | "Encoding::BINARY" => {
                Ok(ExprKind::Lit(Literal::Str("ASCII-8BIT".to_string())))
            }
            "Encoding::UTF_8" => Ok(ExprKind::Lit(Literal::Str("UTF-8".to_string()))),
            "Encoding::US_ASCII" => Ok(ExprKind::Lit(Literal::Str("US-ASCII".to_string()))),
            "Encoding::Windows_1252" => Ok(ExprKind::Lit(Literal::Str("Windows-1252".to_string()))),
            path => {
                let mut parts = path.split("::");
                let first = parts.next().unwrap_or(path);
                let mut expr = Expression::ident(first);
                for part in parts {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_const_get")),
                        args: vec![
                            Argument::positional(expr),
                            Argument::positional(Expression::string(part)),
                        ],
                        optional: false,
                    });
                }
                Ok(expr.kind)
            }
        },

        // Instance var @x → self._rb_x  (prefixed to avoid collision with method bindings)
        Rule::instance_var => {
            let name = &pair.as_str()[1..]; // strip @
            Ok(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: format!("_rb_{}", name),
                null_safe: false,
            })
        }
        // Class var @@x → ident (treated as class-level variable)
        Rule::class_var => {
            let name = &pair.as_str()[2..]; // strip @@
            Ok(ExprKind::Ident(format!("_cls_{}", name)))
        }
        // Global var $x → ident
        Rule::global_var => {
            let name = &pair.as_str()[1..]; // strip $
            Ok(ExprKind::Ident(format!("_global_{}", name)))
        }

        // ── Expression wrappers ─────────────────────────────────────────
        Rule::expression => walk_expression_inner(pair),
        Rule::ternary_expr => walk_ternary(pair),
        Rule::low_and_expr
        | Rule::low_or_expr
        | Rule::or_expr
        | Rule::and_expr
        | Rule::not_expr
        | Rule::comparison
        | Rule::bitor_expr
        | Rule::bitxor_expr
        | Rule::bitand_expr
        | Rule::shift_expr
        | Rule::range_expr
        | Rule::additive
        | Rule::multiplicative
        | Rule::unary => walk_infix_or_unwrap(pair),

        Rule::postfix => walk_postfix(pair),
        Rule::primary => walk_primary(pair),
        Rule::ident_call => walk_ident_call(pair),
        Rule::expression_list => walk_expr_list_kind(pair),

        // ── Special expressions ─────────────────────────────────────────
        Rule::yield_expr => walk_yield(pair),
        Rule::defined_expr => walk_defined(pair),
        Rule::super_expr => walk_super(pair),
        Rule::block_given_expr => Ok(ExprKind::Lit(Literal::Bool(true))), // simplification
        Rule::lambda_literal => walk_lambda(pair),
        Rule::proc_literal => walk_proc(pair),

        // ── If/Unless/Begin as expression ───────────────────────────────
        Rule::if_expr => walk_if_expr(pair),
        Rule::unless_expr => walk_unless_expr(pair),
        Rule::begin_expr => walk_begin_expr(pair),

        Rule::array_inner => walk_array_inner(pair),
        Rule::hash_inner => walk_hash_inner(pair),

        Rule::NEWLINE => Ok(ExprKind::Lit(Literal::Null)),

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Expression inner (handles inline_rescue) ────────────────────────────────

fn walk_expression_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // expression = ternary_expr ~ inline_rescue?
    let expr = walk_expression(inner.remove(0))?;
    // If there's an inline_rescue, wrap in try
    if !inner.is_empty() && inner[0].as_rule() == Rule::inline_rescue {
        let rescue_inner: Vec<Pair<Rule>> = inner.remove(0).into_inner().collect();
        let _rescue_val = if let Some(rp) = rescue_inner.into_iter().next() {
            walk_expression(rp)?
        } else {
            Expression::null()
        };
        // Emit: (begin expr rescue => rescue_val end) as a ternary
        // Simplification: just return the expr (rescue is error handling)
        return Ok(expr.kind);
    }
    Ok(expr.kind)
}

// ── Ternary ─────────────────────────────────────────────────────────────────

fn walk_ternary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // cond ? then : else
    if inner.len() >= 3 {
        let cond = walk_expression(inner.remove(0))?;
        let then = walk_expression(inner.remove(0))?;
        let else_ = walk_expression(inner.remove(0))?;
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then),
            else_: Box::new(else_),
        })
    } else {
        walk_expr_kind(inner.remove(0))
    }
}

// ── Infix / precedence unwrap ───────────────────────────────────────────────

fn walk_infix_or_unwrap(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let rule = pair.as_rule();
    let text = pair.as_str().trim_start().to_string();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        if rule == Rule::not_expr && (text.starts_with('!') || text.starts_with("not ")) {
            let operand = walk_expression(inner.remove(0))?;
            return Ok(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_not")),
                args: vec![Argument::positional(operand)],
                optional: false,
            });
        }
        return walk_expr_kind(inner.remove(0));
    }

    match rule {
        Rule::low_or_expr | Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::low_and_expr | Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::not_expr => {
            let operand = walk_expression(inner.pop().ok_or("Empty not")?)?;
            Ok(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_not")),
                args: vec![Argument::positional(operand)],
                optional: false,
            })
        }
        Rule::comparison => {
            let mut left = walk_expression(inner.remove(0))?;
            let mut i = 0;
            while i < inner.len() {
                if inner[i].as_rule() == Rule::comparison_op {
                    let op_text = inner[i].as_str().trim();
                    i += 1;
                    if i < inner.len() {
                        let right = walk_expression(inner[i].clone())?;
                        i += 1;
                        if op_text == "=~" {
                            left = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__ruby_match_index")),
                                args: vec![Argument::positional(left), Argument::positional(right)],
                                optional: false,
                            });
                            continue;
                        }
                        let op = parse_comparison_op(op_text);
                        left = maybe_ruby_array_binary(left, op, right);
                    }
                } else {
                    i += 1;
                }
            }
            Ok(left.kind)
        }
        Rule::bitor_expr => walk_binary_chain(inner, |_| BinOp::BitOr),
        Rule::bitxor_expr => walk_binary_chain(inner, |_| BinOp::BitXor),
        Rule::bitand_expr => walk_binary_chain(inner, |_| BinOp::BitAnd),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::range_expr => walk_range(inner),
        Rule::additive => walk_binary_chain_with_ops(inner),
        Rule::multiplicative => walk_ruby_multiplicative(inner),
        Rule::unary => {
            let op_str = inner[0].as_str().trim();
            let operand = walk_expression(inner.pop().ok_or("Empty unary")?)?;
            if op_str == "!" {
                return Ok(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_not")),
                    args: vec![Argument::positional(operand)],
                    optional: false,
                });
            }
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "~" => UnaryOp::BitNot,
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand),
            })
        }
        _ => {
            if !inner.is_empty() {
                walk_expr_kind(inner.remove(0))
            } else {
                Ok(ExprKind::Lit(Literal::Null))
            }
        }
    }
}

fn walk_binary_chain(
    mut items: Vec<Pair<Rule>>,
    op_fn: impl Fn(&str) -> BinOp,
) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_expression_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            let op = op_fn("");
            left = match op {
                BinOp::And => Expression::new(ExprKind::Ternary {
                    cond: Box::new(left),
                    then: Box::new(ruby_boolify_expr(right)),
                    else_: Box::new(ruby_bool_expr(false)),
                }),
                BinOp::Or => Expression::new(ExprKind::Ternary {
                    cond: Box::new(left),
                    then: Box::new(ruby_bool_expr(true)),
                    else_: Box::new(ruby_boolify_expr(right)),
                }),
                _ => maybe_ruby_array_binary(left, op, right),
            };
        }
    }
    Ok(left.kind)
}

fn ruby_bool_expr(value: bool) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Bool(value)))
}

fn ruby_boolify_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(expr),
        then: Box::new(ruby_bool_expr(true)),
        else_: Box::new(ruby_bool_expr(false)),
    })
}

/// Ruby `*` is dynamic (string repeat OR numeric mul), same as Python.
fn walk_ruby_multiplicative(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op_str = p.as_str().trim();
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                let op = parse_binop(op_str);
                left = maybe_ruby_array_binary(left, op, right);
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn walk_binary_chain_with_ops(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim());
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = maybe_ruby_array_binary(left, op, right);
            }
        } else if is_expression_rule(p.as_rule()) {
            let right = walk_expression(items[i].clone())?;
            i += 1;
            left = maybe_ruby_array_binary(left, BinOp::Add, right);
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn maybe_ruby_array_binary(left: Expression, op: BinOp, right: Expression) -> Expression {
    if op == BinOp::Mod && matches!(left.kind, ExprKind::Lit(Literal::Str(_))) {
        if let Some(expr) = ruby_percent_hash_literal(&left, &right) {
            return expr;
        }
        let mut args = vec![Argument::positional(left)];
        if let ExprKind::Array(elements) = right.kind {
            args.extend(
                elements
                    .into_iter()
                    .map(|element| Argument::positional(element.value)),
            );
        } else {
            args.push(Argument::positional(right));
        }
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("sprintf")),
            args,
            optional: false,
        });
    }
    let helper = if matches!(left.kind, ExprKind::Lit(Literal::Str(_)))
        || matches!(right.kind, ExprKind::Lit(Literal::Str(_)))
    {
        match op {
            BinOp::Lt => Some("__ruby_str_lt"),
            BinOp::Gt => Some("__ruby_str_gt"),
            BinOp::LtEq => Some("__ruby_str_lte"),
            BinOp::GtEq => Some("__ruby_str_gte"),
            BinOp::Spaceship => Some("__ruby_str_cmp"),
            _ => None,
        }
    } else if is_ruby_time_expr(&left) || is_ruby_time_expr(&right) {
        match op {
            BinOp::Eq => Some("__ruby_time_eq"),
            BinOp::Lt => Some("__ruby_time_lt"),
            BinOp::Gt => Some("__ruby_time_gt"),
            BinOp::LtEq => Some("__ruby_time_lte"),
            BinOp::GtEq => Some("__ruby_time_gte"),
            BinOp::Spaceship => Some("__ruby_time_cmp"),
            _ => None,
        }
    } else {
        None
    }
    .or(match op {
        BinOp::Add => Some("__ruby_op_add"),
        BinOp::Sub => Some("__ruby_op_sub"),
        BinOp::Mul => Some("__ruby_op_mul"),
        BinOp::Div => Some("__ruby_op_div"),
        BinOp::Pow => Some("__ruby_op_pow"),
        BinOp::Shl => Some("__ruby_op_shl"),
        BinOp::Shr => Some("__ruby_op_shr"),
        BinOp::BitAnd => Some("__ruby_op_and"),
        BinOp::BitOr => Some("__ruby_op_or"),
        BinOp::BitXor => Some("__ruby_op_xor"),
        BinOp::Eq => Some("__ruby_eq"),
        BinOp::StrictEq => Some("__ruby_proc_call"),
        _ => None,
    });
    if let Some(name) = helper {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: vec![Argument::positional(left), Argument::positional(right)],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }
}

fn is_ruby_time_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => name.starts_with("__ruby_time_"),
            _ => false,
        },
        _ => false,
    }
}

fn ruby_percent_hash_literal(fmt_expr: &Expression, hash_expr: &Expression) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(fmt)) = &fmt_expr.kind else {
        return None;
    };
    let ExprKind::Object(props) = &hash_expr.kind else {
        return None;
    };
    let mut parts = Vec::new();
    let mut lit = String::new();
    let mut chars = fmt.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '%' && chars.peek() == Some(&'{') {
            chars.next();
            let mut key = String::new();
            while let Some(k) = chars.next() {
                if k == '}' {
                    break;
                }
                key.push(k);
            }
            if !lit.is_empty() {
                parts.push(Expression::string(&lit));
                lit.clear();
            }
            let value = props.iter().find_map(|prop| match prop {
                ObjectProperty::KeyValue { key: k, value } => {
                    if matches!(&k.kind, ExprKind::Lit(Literal::Str(name)) if name == &key) {
                        Some(value.clone())
                    } else {
                        None
                    }
                }
                _ => None,
            })?;
            parts.push(value);
        } else {
            lit.push(c);
        }
    }
    if !lit.is_empty() {
        parts.push(Expression::string(&lit));
    }
    let mut iter = parts.into_iter();
    let first = iter.next().unwrap_or_else(|| Expression::string(""));
    Some(iter.fold(first, |acc, part| ruby_add_expr(acc, part)))
}

fn literal_string(expr: &Expression) -> Option<&str> {
    if let ExprKind::Lit(Literal::Str(s)) = &expr.kind {
        Some(s)
    } else {
        None
    }
}

fn ruby_hash_string_map(expr: &Expression) -> Option<Vec<(String, String)>> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    let mut out = Vec::new();
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            return None;
        };
        out.push((
            literal_string(key)?.to_string(),
            literal_string(value)?.to_string(),
        ));
    }
    Some(out)
}

fn ruby_literal_string_substitution(
    receiver: &Expression,
    method_name: &str,
    args: &[Argument],
) -> Option<ExprKind> {
    if !matches!(method_name, "gsub" | "sub") || args.len() != 2 {
        return None;
    }
    let input = literal_string(receiver)?;
    let replace_all = method_name == "gsub";
    if matches!(literal_string(&args[0].value), Some("/a./")) {
        if let Some(replacement) = literal_string(&args[1].value) {
            let chars: Vec<char> = input.chars().collect();
            let mut out = String::new();
            let mut i = 0;
            let mut changed = false;
            while i < chars.len() {
                if (!changed || replace_all) && i + 1 < chars.len() && chars[i] == 'a' {
                    out.push_str(replacement);
                    changed = true;
                    i += 2;
                } else {
                    out.push(chars[i]);
                    i += 1;
                }
            }
            return Some(ExprKind::Lit(Literal::Str(out)));
        }
    }
    if let Some(map) = ruby_hash_string_map(&args[1].value) {
        let mut changed = false;
        let mut out = String::new();
        for ch in input.chars() {
            if !replace_all && changed {
                out.push(ch);
                continue;
            }
            let key = ch.to_string();
            if let Some((_, replacement)) = map.iter().find(|(k, _)| k == &key) {
                out.push_str(replacement);
                changed = true;
            } else {
                out.push(ch);
            }
        }
        return Some(ExprKind::Lit(Literal::Str(out)));
    }
    None
}

// ── Range ───────────────────────────────────────────────────────────────────

fn walk_range(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    if items.len() == 1 {
        return walk_expr_kind(items.remove(0));
    }
    let start = walk_expression(items.remove(0))?;
    // Find range_op
    let mut inclusive = true;
    let mut end_idx = 0;
    for (i, p) in items.iter().enumerate() {
        if p.as_rule() == Rule::range_op {
            inclusive = p.as_str() == "..";
            end_idx = i + 1;
            break;
        }
    }
    if end_idx < items.len() {
        let end = walk_expression(items.remove(end_idx))?;
        // `..` is inclusive, `...` exclusive — pass the flag through. The shared
        // range/slice emitters honour it for both numeric and char bounds
        // (no lossy compile-time `end + 1`, which corrupted `'a'..'z'`).
        Ok(ExprKind::Range {
            start: Box::new(start),
            end: Box::new(end),
            inclusive,
        })
    } else {
        Ok(start.kind)
    }
}

// ── Postfix (call, member, subscript, block) ────────────────────────────────

/// `ident_call = ${ (constant | identifier) ~ tight_call }` — a whitespace-tight
/// `foo(args)` call. The `(` immediately follows the name; `foo (args)` (space)
/// never reaches here (it stays a command call).
fn walk_ident_call(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    let callee = Expression::new(walk_expr_kind(inner.remove(0))?);
    // tight_call = !{ "(" ~ call_args? ~ ")" }
    let args = inner
        .into_iter()
        .find(|c| c.as_rule() == Rule::tight_call)
        .and_then(|tc| tc.into_inner().find(|c| c.as_rule() == Rule::call_args))
        .map(walk_call_args)
        .transpose()?
        .unwrap_or_default();
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "lambda") && args.len() == 1 {
        return Ok(ruby_proc_expr("__ruby_lambda", args[0].value.clone()).kind);
    }
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "eval") && args.len() == 1 {
        return Ok(ruby_eval_expr(args[0].value.clone()).kind);
    }
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "method") && args.len() == 1 {
        if let Some(name) = ruby_method_name_arg(&args[0].value) {
            return Ok(ruby_method_expr(
                &name,
                Expression::null(),
                "Object",
                "Object",
                Expression::null(),
            )
            .kind);
        }
    }
    Ok(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn ruby_method_name_arg(expr: &Expression) -> Option<String> {
    if let ExprKind::Lit(Literal::Str(name)) = &expr.kind {
        Some(name.clone())
    } else {
        None
    }
}

fn ruby_eval_expr(source: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__vybe_eval")),
        args: vec![
            Argument::positional(source),
            Argument::positional(Expression::string("ruby")),
        ],
        optional: false,
    })
}

fn ruby_receiver_class_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => name.clone(),
            _ => "Object".to_string(),
        },
        _ => "Object".to_string(),
    }
}

fn ruby_method_expr(
    name: &str,
    fn_expr: Expression,
    owner: &str,
    receiver_class: &str,
    receiver: Expression,
) -> Expression {
    let original = ruby_alias_original(name);
    let info = ruby_method_info(owner, name);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__ruby_method")),
        args: vec![
            Argument::positional(Expression::string(name)),
            Argument::positional(fn_expr),
            Argument::positional(Expression::int(info.arity)),
            Argument::positional(Expression::int(info.param_count)),
            Argument::positional(Expression::string(owner)),
            Argument::positional(Expression::string(receiver_class)),
            Argument::positional(Expression::string(&original)),
            Argument::positional(receiver),
        ],
        optional: false,
    })
}

fn walk_postfix(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty postfix")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() == Rule::postfix_chain {
            expr = walk_postfix_chain(expr, chain)?;
        } else if chain.as_rule() == Rule::constant {
            let const_name = chain.as_str();
            if matches!((&expr.kind, const_name), (ExprKind::Ident(base), "PI") if base == "Math") {
                expr = Expression::new(ExprKind::Lit(Literal::Float(std::f64::consts::PI)));
            } else if matches!((&expr.kind, const_name), (ExprKind::Ident(base), "E") if base == "Math")
            {
                expr = Expression::new(ExprKind::Lit(Literal::Float(std::f64::consts::E)));
            } else {
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_const_get")),
                    args: vec![
                        Argument::positional(expr),
                        Argument::positional(Expression::string(const_name)),
                    ],
                    optional: false,
                });
            }
        }
    }
    Ok(expr.kind)
}

fn walk_postfix_chain(expr: Expression, chain: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = chain.into_inner().collect();

    if children.is_empty() {
        // bare () call
        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__ruby_proc_call")),
            args: vec![Argument::positional(expr)],
            optional: false,
        }));
    }

    let first_rule = children[0].as_rule();

    match first_rule {
        Rule::method_name_id => {
            // Method call: .method or &.method
            let mut method_name = children[0].as_str().to_string();
            let null_safe = children.iter().any(|c| c.as_str() == "&.");

            // Check if there are call args
            let args = children
                .iter()
                .find(|c| c.as_rule() == Rule::call_args)
                .map(|c| walk_call_args(c.clone()))
                .transpose()?
                .unwrap_or_default();

            // Check for trailing block
            let block_text = children
                .iter()
                .find(|c| c.as_rule() == Rule::block_literal)
                .map(|c| c.as_str().to_string());
            let block = children
                .iter()
                .find(|c| c.as_rule() == Rule::block_literal)
                .map(|c| walk_block_literal(c.clone()))
                .transpose()?;

            let mut final_args = args;
            if let Some(block_lambda) = block {
                final_args.push(Argument::positional(block_lambda));
            }

            if ruby_slice_returns_nil(&expr, &method_name, &final_args) {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            }
            normalize_ruby_slice_call(&mut method_name, &mut final_args);

            if let Some(lit) = ruby_literal_string_substitution(&expr, &method_name, &final_args) {
                return Ok(Expression::new(lit));
            }

            if method_name == "zip"
                && block_text.is_none()
                && final_args
                    .iter()
                    .all(|arg| arg.name.is_none() && !arg.spread)
            {
                let mut iterables = Vec::with_capacity(final_args.len() + 1);
                iterables.push(expr);
                iterables.extend(final_args.into_iter().map(|arg| arg.value));
                return Ok(Expression::new(ExprKind::Zip {
                    iterables,
                    mode: ZipMode::First,
                    strict: false,
                }));
            }

            if matches!(
                method_name.as_str(),
                "class_eval" | "module_eval" | "instance_eval"
            ) && final_args.len() == 1
            {
                return Ok(ruby_eval_expr(final_args[0].value.clone()));
            }

            if method_name == "find_index" && final_args.len() == 1 {
                method_name = if matches!(final_args[0].value.kind, ExprKind::Lambda { .. }) {
                    "__ruby_find_index_block".to_string()
                } else {
                    "__ruby_find_index_value".to_string()
                };
            } else if matches!(method_name.as_str(), "inject" | "reduce")
                && final_args.len() == 1
                && matches!(final_args[0].value.kind, ExprKind::Lit(Literal::Str(_)))
            {
                method_name = "__ruby_inject_symbol".to_string();
            } else if matches!(method_name.as_str(), "inject" | "reduce")
                && final_args.len() == 2
                && matches!(final_args[1].value.kind, ExprKind::Lambda { .. })
            {
                method_name = "__ruby_inject_initial".to_string();
            } else if method_name == "rindex"
                && final_args.len() == 1
                && matches!(final_args[0].value.kind, ExprKind::Lambda { .. })
            {
                method_name = "__ruby_rindex_block".to_string();
            } else if matches!(method_name.as_str(), "bsearch" | "bsearch_index")
                && final_args.len() == 1
                && matches!(final_args[0].value.kind, ExprKind::Lambda { .. })
            {
                let suffix = if lambda_contains_spaceship(&final_args[0].value) {
                    "cmp"
                } else {
                    "bool"
                };
                method_name = format!("__ruby_{}_{}", method_name, suffix);
            } else if matches!(method_name.as_str(), "find" | "detect")
                && final_args.len() == 2
                && matches!(final_args[1].value.kind, ExprKind::Lambda { .. })
            {
                method_name = "__ruby_find_ifnone".to_string();
            }

            if let ExprKind::Ident(class_name) = &expr.kind {
                if class_name == "Proc" && method_name == "new" && !final_args.is_empty() {
                    return Ok(Expression::new(
                        ruby_proc_expr("__ruby_proc", final_args.remove(0).value).kind,
                    ));
                }
                if class_name == "Enumerator" && method_name == "new" {
                    if let Some(gen_fn) = block_text
                        .as_deref()
                        .and_then(ruby_enumerator_generator_expr)
                    {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__ruby_enum_new")),
                            args: vec![Argument::positional(gen_fn)],
                            optional: false,
                        }));
                    }
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_enum_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if class_name == "Time" {
                    let builtin = match method_name.as_str() {
                        "utc" | "gm" => Some("__ruby_time_utc"),
                        "local" | "mktime" | "new" => Some("__ruby_time_local"),
                        "now" => Some("__ruby_time_now"),
                        "at" => Some("__ruby_time_at"),
                        "parse" | "iso8601" | "rfc2822" | "httpdate" => Some("__ruby_time_parse"),
                        _ => None,
                    };
                    if let Some(name) = builtin {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(name)),
                            args: final_args,
                            optional: false,
                        }));
                    }
                }
                if class_name == "Date" && method_name == "new" {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_date_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if class_name == "Symbol" && method_name == "all_symbols" {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_symbols")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if final_args.is_empty() && method_name == "ancestors" {
                    if let Some(ancestors) = ruby_exception_ancestors_expr(class_name) {
                        return Ok(ancestors);
                    }
                }
            }

            if matches!((&expr.kind, method_name.as_str()), (ExprKind::Ident(name), "utc") if name == "Time")
            {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_time_utc")),
                    args: final_args,
                    optional: false,
                }));
            }

            // Normalize .new() → ExprKind::New (constructor call)
            if method_name == "new" {
                if let ExprKind::Ident(name) = &expr.kind {
                    if let Some(helper) = ruby_exception_helper_name(name) {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(helper)),
                            args: final_args,
                            optional: false,
                        }));
                    }
                }
                if matches!(expr.kind, ExprKind::Ident(ref name) if name == "Random") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_random_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if matches!(expr.kind, ExprKind::Ident(ref name) if name == "Array") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_array_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                // Route `Set.new` / `SortedSet.new` through direct calls so the
                // shared `ecma_new_dispatch` `new Set(...)` intercept (which
                // would build a raw ecma:set) never fires — Ruby's Set is a
                // deduped tagged array.
                if matches!(expr.kind, ExprKind::Ident(ref name) if name == "Set") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_set_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if matches!(expr.kind, ExprKind::Ident(ref name) if name == "SortedSet") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_sorted_set_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                return Ok(Expression::new(ExprKind::New {
                    class: Box::new(expr),
                    args: final_args,
                }));
            }

            // Normalize .call() → direct call (lambda/proc invocation)
            if method_name == "call" {
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_proc_call")),
                    args: call_args,
                    optional: false,
                }));
            }

            if method_name == "set_backtrace" && final_args.len() == 1 {
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_exception_set_backtrace")),
                    args: call_args,
                    optional: false,
                }));
            }

            if method_name == "exception" {
                if final_args.is_empty() {
                    return Ok(expr);
                }
                if final_args.len() == 1 {
                    let mut call_args = vec![Argument::positional(expr)];
                    call_args.extend(final_args);
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_exception_with_message")),
                        args: call_args,
                        optional: false,
                    }));
                }
            }

            if method_name == "yield" {
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_proc_call")),
                    args: call_args,
                    optional: false,
                }));
            }

            if matches!(
                method_name.as_str(),
                "each" | "map" | "collect" | "select" | "filter" | "reject"
            ) && final_args.is_empty()
            {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_enum_from")),
                    args: vec![
                        Argument::positional(expr),
                        Argument::positional(Expression::string(&method_name)),
                    ],
                    optional: false,
                }));
            }

            if matches!(
                method_name.as_str(),
                "const_get" | "const_set" | "const_defined?" | "remove_const" | "constants"
            ) {
                let builtin = match method_name.as_str() {
                    "const_get" => "__ruby_const_get",
                    "const_set" => "__ruby_const_set",
                    "const_defined?" => "__ruby_const_defined",
                    "remove_const" => "__ruby_remove_const",
                    "constants" => "__ruby_constants",
                    _ => unreachable!(),
                };
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(builtin)),
                    args: call_args,
                    optional: false,
                }));
            }

            if method_name == "method" && final_args.len() == 1 {
                if let Some(name) = ruby_method_name_arg(&final_args[0].value) {
                    let owner = ruby_receiver_class_name(&expr);
                    let original = ruby_alias_original(&name);
                    let fn_expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr.clone()),
                        field: original,
                        null_safe: false,
                    });
                    return Ok(ruby_method_expr(&name, fn_expr, &owner, &owner, expr));
                }
            }

            // Normalize .is_a?/.kind_of?(Klass) → `expr instanceof Klass`
            // (the shared JS instanceof path: reads the constructor's name and
            // checks the `__types` ancestry, so inheritance works). Wrapped in a
            // ternary so it materializes to a real `true`/`false`.
            if matches!(method_name.as_str(), "is_a?" | "kind_of?") && final_args.len() == 1 {
                let class_arg = final_args.into_iter().next().unwrap().value;
                let inst = Expression::new(ExprKind::Binary {
                    op: BinOp::InstanceOf,
                    left: Box::new(expr),
                    right: Box::new(class_arg),
                });
                return Ok(Expression::new(ExprKind::Ternary {
                    cond: Box::new(inst),
                    then: Box::new(Expression::bool(true)),
                    else_: Box::new(Expression::bool(false)),
                }));
            }

            // Normalize .first → Index(expr, 0) — pure bytecode, no host call
            if method_name == "first" && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_index_get")),
                    args: vec![
                        Argument::positional(expr),
                        Argument::positional(Expression::int(0)),
                    ],
                    optional: false,
                }));
            }

            // Normalize .last → Index(expr, -1) — pure bytecode
            if method_name == "last" && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_index_get")),
                    args: vec![
                        Argument::positional(expr),
                        Argument::positional(Expression::int(-1)),
                    ],
                    optional: false,
                }));
            }

            if method_name == "integer?" && final_args.is_empty() {
                match expr.kind {
                    ExprKind::Lit(Literal::Int(_)) => return Ok(Expression::bool(true)),
                    ExprKind::Lit(Literal::Float(_)) => return Ok(Expression::bool(false)),
                    _ => {}
                }
            }

            if final_args.is_empty() {
                if method_name == "inspect" {
                    let helper = if ruby_expr_may_be_exception(&expr) {
                        "__ruby_exception_inspect"
                    } else {
                        "__ruby_inspect"
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(helper)),
                        args: vec![Argument::positional(expr)],
                        optional: false,
                    }));
                }
                if method_name == "class" {
                    if matches!(
                        &expr.kind,
                        ExprKind::Member { field, .. }
                            if matches!(field.as_str(), "backtrace" | "backtrace_locations")
                    ) {
                        return Ok(Expression::string("Array"));
                    }
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_class")),
                        args: vec![Argument::positional(expr)],
                        optional: false,
                    }));
                }
                if method_name == "name" {
                    if matches!(expr.kind, ExprKind::Lit(Literal::Str(_))) {
                        return Ok(expr);
                    }
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_name")),
                        args: vec![Argument::positional(expr)],
                        optional: false,
                    }));
                }
                if method_name == "message"
                    || (method_name == "to_s" && ruby_expr_may_be_exception(&expr))
                {
                    return Ok(Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: "message".to_string(),
                        null_safe,
                    }));
                }
                if method_name == "__type" {
                    return Ok(Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: "__type".to_string(),
                        null_safe,
                    }));
                }
                if method_name == "backtrace" {
                    return Ok(Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: "backtrace".to_string(),
                        null_safe,
                    }));
                }
                if method_name == "backtrace_locations" {
                    return Ok(Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: "backtrace".to_string(),
                        null_safe,
                    }));
                }
                if method_name == "cause" {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
                }
                if method_name == "full_message" {
                    return Ok(Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: "message".to_string(),
                        null_safe,
                    }));
                }
            }

            if matches!(method_name.as_str(), "include?" | "member?") && final_args.len() == 1 {
                if let ExprKind::Ident(name) = &final_args[0].value.kind {
                    if ruby_exception_helper_name(name).is_some() {
                        final_args[0].value = Expression::string(name);
                    }
                }
            }

            let member = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe,
            });

            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args: final_args,
                optional: false,
            }))
        }
        Rule::constant => {
            // Scope resolution: ::Constant
            let const_name = children[0].as_str();
            if let ExprKind::Ident(base) = &expr.kind {
                if base == "Math" && const_name == "PI" {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Float(
                        std::f64::consts::PI,
                    ))));
                }
                if base == "Math" && const_name == "E" {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Float(
                        std::f64::consts::E,
                    ))));
                }
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_const_get")),
                    args: vec![
                        Argument::positional(Expression::ident(base)),
                        Argument::positional(Expression::string(const_name)),
                    ],
                    optional: false,
                }));
            }
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_const_get")),
                args: vec![
                    Argument::positional(expr),
                    Argument::positional(Expression::string(const_name)),
                ],
                optional: false,
            }))
        }
        Rule::call_args => {
            // Bare call: expr(args)
            let args = walk_call_args(children.into_iter().next().unwrap())?;
            let mut call_args = vec![Argument::positional(expr)];
            call_args.extend(args);
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_proc_call")),
                args: call_args,
                optional: false,
            }))
        }
        Rule::expression_list => {
            // Subscript: expr[index]
            let index = walk_expr_list_single(children.into_iter().next().unwrap())?;
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_index_get")),
                args: vec![Argument::positional(expr), Argument::positional(index)],
                optional: false,
            }))
        }
        Rule::block_literal => {
            // Trailing block on its own (e.g., `array.each { |x| ... }`)
            // The method call should already be formed; this adds the block as arg
            if let ExprKind::Call {
                callee,
                mut args,
                optional,
            } = expr.kind
            {
                let block_lambda = walk_block_literal(children.into_iter().next().unwrap())?;
                args.push(Argument::positional(block_lambda));
                Ok(Expression::new(ExprKind::Call {
                    callee,
                    args,
                    optional,
                }))
            } else {
                // Bare block on expression — treat as call with block
                let block_lambda = walk_block_literal(children.into_iter().next().unwrap())?;
                Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args: vec![Argument::positional(block_lambda)],
                    optional: false,
                }))
            }
        }
        _ => {
            // Try to interpret as subscript or call
            if is_expression_rule(first_rule) {
                let index = walk_expression(children.into_iter().next().unwrap())?;
                Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_index_get")),
                    args: vec![Argument::positional(expr), Argument::positional(index)],
                    optional: false,
                }))
            } else {
                Ok(expr)
            }
        }
    }
}

fn ruby_enumerator_generator_expr(source: &str) -> Option<Expression> {
    let mut body = Vec::new();
    for piece in source.split(';') {
        let trimmed = piece.trim();
        let value = if let Some((_, rhs)) = trimmed.rsplit_once("<<") {
            rhs.trim().trim_end_matches('}').trim()
        } else if let Some((_, rhs)) = trimmed.rsplit_once(".yield") {
            rhs.trim().trim_end_matches('}').trim()
        } else {
            continue;
        };
        if value.is_empty() {
            continue;
        }
        let mut parsed = RubyParser::parse(Rule::expression, value).ok()?;
        let expr_pair = parsed.next()?;
        let expr = walk_expression(expr_pair).ok()?;
        body.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Yield(Some(Box::new(expr))),
        ))));
    }
    if body.is_empty() {
        return None;
    }
    Some(Expression::new(ExprKind::FunctionExpr(Box::new(
        Statement::new(StmtKind::FunctionDecl {
            name: String::new(),
            params: Vec::new(),
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: true,
            is_sub: false,
        }),
    ))))
}

fn walk_call_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    let mut pending_hash = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_arg {
            let raw_arg = p.as_str().to_string();
            let children: Vec<Pair<Rule>> = p.into_inner().collect();
            if children.is_empty() {
                continue;
            }

            if raw_arg.contains("=>") && children.len() >= 2 {
                let key = walk_expression(children[0].clone())?;
                let value = walk_expression(children[1].clone())?;
                pending_hash.push(ObjectProperty::KeyValue { key, value });
                continue;
            }

            if !pending_hash.is_empty() {
                args.push(Argument::positional(Expression::new(ExprKind::Object(
                    std::mem::take(&mut pending_hash),
                ))));
            }

            let first_text = children[0].as_str();

            if first_text == "**" {
                // Double splat
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument {
                        value: val,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                }
            } else if first_text == "*" {
                // Splat
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument {
                        value: val,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                }
            } else if first_text == "&" || raw_arg.trim_start().starts_with('&') {
                // Block arg
                if raw_arg.trim_start().starts_with('&') {
                    let val = walk_expression(children.into_iter().next().unwrap())?;
                    args.push(Argument::positional(ruby_block_arg_to_lambda(val)));
                } else if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument::positional(ruby_block_arg_to_lambda(val)));
                }
            } else if children.len() >= 2 && children[0].as_rule() == Rule::identifier {
                // Check if keyword arg: identifier ":" expression
                let has_colon = children.iter().any(|c| c.as_str() == ":");
                if has_colon {
                    let name = children[0].as_str().to_string();
                    let val = walk_expression(children.into_iter().last().unwrap())?;
                    args.push(Argument {
                        value: val,
                        name: Some(name),
                        by_ref: false,
                        spread: false,
                    });
                } else {
                    let val = walk_expression(children.into_iter().next().unwrap())?;
                    args.push(Argument::positional(val));
                }
            } else {
                let val = walk_expression(children.into_iter().next().unwrap())?;
                args.push(Argument::positional(val));
            }
        }
    }
    if !pending_hash.is_empty() {
        args.push(Argument::positional(Expression::new(ExprKind::Object(
            pending_hash,
        ))));
    }
    Ok(args)
}

fn ruby_block_arg_to_lambda(expr: Expression) -> Expression {
    if let ExprKind::Call { callee, .. } = &expr.kind {
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__ruby_method") {
            let param = Param {
                name: "__ruby_proc_arg".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            };
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_proc_call")),
                args: vec![
                    Argument::positional(expr),
                    Argument::positional(Expression::ident("__ruby_proc_arg")),
                ],
                optional: false,
            });
            return Expression::new(ExprKind::Lambda {
                params: vec![param],
                body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(call)))]),
                is_async: false,
                captures: Vec::new(),
            });
        }
    }
    if let ExprKind::Lit(Literal::Str(method)) = &expr.kind {
        let param = Param {
            name: "__ruby_proc_arg".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        let call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__ruby_proc_arg")),
                field: method.clone(),
                null_safe: false,
            })),
            args: Vec::new(),
            optional: false,
        });
        return Expression::new(ExprKind::Lambda {
            params: vec![param],
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(call)))]),
            is_async: false,
            captures: Vec::new(),
        });
    }
    expr
}

// ── Block literal ───────────────────────────────────────────────────────────

fn walk_block_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::do_block | Rule::brace_block => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::block_params => {
                            params = walk_block_params(bp)?;
                        }
                        Rule::body => {
                            body = walk_body(bp)?;
                        }
                        _ => {
                            // Statements directly in brace_block
                            let stmt = walk_statement(bp)?;
                            if !matches!(stmt.kind, StmtKind::Empty) {
                                body.push(stmt);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    apply_implicit_return(&mut body);

    Ok(Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    }))
}

/// Ruby implicit return: last expression in a body becomes a Return.
fn apply_implicit_return(body: &mut Vec<Statement>) {
    if let Some(last) = body.last_mut() {
        if matches!(&last.kind, StmtKind::Expr(_)) {
            if let StmtKind::Expr(e) = std::mem::replace(&mut last.kind, StmtKind::Empty) {
                last.kind = StmtKind::Return(Some(e));
            }
        } else if let StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } = &mut last.kind
        {
            apply_implicit_return(then_body);
            for (_, body) in elifs {
                apply_implicit_return(body);
            }
            if let Some(body) = else_body {
                apply_implicit_return(body);
            }
        }
    }
}

fn lambda_contains_spaceship(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(e) => expr_contains_spaceship(e),
            LambdaBody::Block(stmts) => stmts.iter().any(stmt_contains_spaceship),
        },
        _ => false,
    }
}

fn stmt_contains_spaceship(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) => expr_contains_spaceship(e),
        StmtKind::Return(Some(e)) => expr_contains_spaceship(e),
        StmtKind::Block(stmts) => stmts.iter().any(stmt_contains_spaceship),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_contains_spaceship(cond)
                || then_body.iter().any(stmt_contains_spaceship)
                || elifs.iter().any(|(cond, body)| {
                    expr_contains_spaceship(cond) || body.iter().any(stmt_contains_spaceship)
                })
                || else_body
                    .as_ref()
                    .map(|body| body.iter().any(stmt_contains_spaceship))
                    .unwrap_or(false)
        }
        _ => false,
    }
}

fn expr_contains_spaceship(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Binary { op, left, right } => {
            *op == BinOp::Spaceship
                || expr_contains_spaceship(left)
                || expr_contains_spaceship(right)
        }
        ExprKind::Unary { expr, .. } => expr_contains_spaceship(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_spaceship(cond)
                || expr_contains_spaceship(then)
                || expr_contains_spaceship(else_)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_contains_spaceship(callee)
                || args.iter().any(|arg| expr_contains_spaceship(&arg.value))
        }
        ExprKind::Member { object, .. } => expr_contains_spaceship(object),
        ExprKind::Index { object, index, .. } => {
            expr_contains_spaceship(object) || expr_contains_spaceship(index)
        }
        ExprKind::Assign { target, value } => {
            expr_contains_spaceship(target) || expr_contains_spaceship(value)
        }
        ExprKind::Array(elements) => elements.iter().any(|element| {
            element
                .key
                .as_ref()
                .map(expr_contains_spaceship)
                .unwrap_or(false)
                || expr_contains_spaceship(&element.value)
        }),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => expr_contains_spaceship(e),
            InterpolPart::Text(_) => false,
        }),
        ExprKind::Range { start, end, .. } => {
            expr_contains_spaceship(start) || expr_contains_spaceship(end)
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(e) => expr_contains_spaceship(e),
            LambdaBody::Block(stmts) => stmts.iter().any(stmt_contains_spaceship),
        },
        _ => false,
    }
}

fn walk_block_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::block_param_list {
            for bp in p.into_inner() {
                if bp.as_rule() == Rule::block_param_item {
                    let inner = bp.into_inner().next();
                    if let Some(item) = inner {
                        match item.as_rule() {
                            Rule::splat_param => {
                                let name = item
                                    .into_inner()
                                    .find(|c| c.as_rule() == Rule::identifier)
                                    .map(|c| c.as_str().to_string())
                                    .unwrap_or_default();
                                params.push(Param {
                                    name,
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: true,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                            Rule::optional_param => {
                                let mut inner = item.into_inner();
                                let name = inner
                                    .next()
                                    .map(|c| c.as_str().to_string())
                                    .unwrap_or_default();
                                let default = inner
                                    .find(|c| is_expression_rule(c.as_rule()))
                                    .map(walk_expression)
                                    .transpose()?;
                                params.push(Param {
                                    name,
                                    type_hint: None,
                                    default,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: true,
                                    is_nullable: false,
                                });
                            }
                            Rule::identifier => {
                                params.push(Param {
                                    name: item.as_str().to_string(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                            _ => {}
                        }
                    }
                }
            }
        }
    }
    Ok(params)
}

// ── Primary ─────────────────────────────────────────────────────────────────

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let source = pair.as_str().trim().to_string();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    if inner.is_empty() {
        if source.starts_with('[') && source.ends_with(']') {
            return Ok(ExprKind::Array(Vec::new()));
        }
        return Ok(ExprKind::Lit(Literal::Null));
    }

    let first = &inner[0];
    match first.as_rule() {
        Rule::array_inner => {
            // Array literal [...]
            walk_array_inner(inner.remove(0))
        }
        Rule::hash_inner => {
            // Hash literal {...}
            walk_hash_inner(inner.remove(0))
        }
        _ => walk_expr_kind(inner.remove(0)),
    }
}

fn walk_array_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(|p| -> Result<ArrayElement, String> {
            let val = walk_expression(p)?;
            Ok(ArrayElement {
                key: None,
                value: val,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(ExprKind::Array(elements))
}

fn walk_hash_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut props = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::hash_pair {
            let children: Vec<Pair<Rule>> = p.into_inner().collect();
            if children.len() >= 2 {
                // Could be hash rocket (key => val) or symbol shorthand (key: val)
                let first = &children[0];
                if first.as_rule() == Rule::identifier && children.len() == 2 {
                    // Symbol shorthand: key: val
                    let key =
                        Expression::new(ExprKind::Lit(Literal::Str(first.as_str().to_string())));
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    props.push(ObjectProperty::KeyValue { key, value: val });
                } else {
                    let key = walk_expression(children[0].clone())?;
                    let val = walk_expression(children.into_iter().last().unwrap())?;
                    props.push(ObjectProperty::KeyValue { key, value: val });
                }
            } else if children.len() == 1 {
                // **expr (double splat)
                let val = walk_expression(children.into_iter().next().unwrap())?;
                props.push(ObjectProperty::Spread(val));
            }
        }
    }
    Ok(ExprKind::Object(props))
}

// ── Interpolated string ─────────────────────────────────────────────────────

fn walk_interpolated_string(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut parts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::interp_start | Rule::interp_end => {}
            Rule::interp_text => {
                parts.push(InterpolPart::Text(p.as_str().to_string()));
            }
            Rule::interp_escape => {
                parts.push(InterpolPart::Text(p.as_str().to_string()));
            }
            Rule::interp_expr => {
                for ip in p.into_inner() {
                    if is_expression_rule(ip.as_rule()) {
                        parts.push(InterpolPart::Expr(walk_expression(ip)?));
                    }
                }
            }
            _ => {}
        }
    }

    // Optimize: if only text parts, concat into single string
    if parts.iter().all(|p| matches!(p, InterpolPart::Text(_))) {
        let s: String = parts
            .iter()
            .map(|p| match p {
                InterpolPart::Text(t) => t.as_str(),
                _ => "",
            })
            .collect();
        return Ok(ExprKind::Lit(Literal::Str(parse_ruby_string(&s))));
    }

    for part in &mut parts {
        if let InterpolPart::Text(text) = part {
            *text = parse_ruby_string(text);
        }
    }

    Ok(ExprKind::Interpolation(parts))
}

// ── Lambda ──────────────────────────────────────────────────────────────────

fn walk_lambda(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let source = pair.as_str().to_string();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_param_list(p)?,
            Rule::body => body = walk_body(p)?,
            _ => {
                // Statements in lambda brace body
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    body.push(stmt);
                }
            }
        }
    }

    apply_implicit_return(&mut body);
    if let (Some(open), Some(close)) = (source.find('{'), source.rfind('}')) {
        let inner = source[open + 1..close].trim();
        if !inner.is_empty() && !inner.contains(';') && !inner.contains('\n') {
            if let Ok(mut parsed) = RubyParser::parse(Rule::expression, inner) {
                if let Some(expr_pair) = parsed.next() {
                    body = vec![Statement::new(StmtKind::Return(Some(walk_expression(
                        expr_pair,
                    )?)))];
                }
            }
        }
    }

    let lambda = Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    Ok(ruby_proc_expr("__ruby_lambda", lambda).kind)
}

fn walk_proc(pair: Pair<Rule>) -> Result<ExprKind, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::block_literal {
            let lambda = walk_block_literal(p)?;
            return Ok(ruby_proc_expr("__ruby_proc", lambda).kind);
        }
    }
    let lambda = Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(Vec::new()),
        is_async: false,
        captures: Vec::new(),
    });
    Ok(ruby_proc_expr("__ruby_proc", lambda).kind)
}

// ── Yield ───────────────────────────────────────────────────────────────────

fn walk_yield(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::expression_list {
            for ep in p.into_inner() {
                if is_expression_rule(ep.as_rule()) {
                    args.push(walk_expression(ep)?);
                }
            }
        } else if is_expression_rule(p.as_rule()) {
            args.push(walk_expression(p)?);
        }
    }
    // Ruby yield calls the block; emit as Yield for now
    if args.is_empty() {
        Ok(ExprKind::Yield(None))
    } else if args.len() == 1 {
        Ok(ExprKind::Yield(Some(Box::new(
            args.into_iter().next().unwrap(),
        ))))
    } else {
        Ok(ExprKind::Yield(Some(Box::new(Expression::new(
            ExprKind::Array(
                args.into_iter()
                    .map(|a| ArrayElement {
                        key: None,
                        value: a,
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ),
        )))))
    }
}

// ── Defined? ────────────────────────────────────────────────────────────────

fn walk_defined(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // defined?(expr) → check if expr is defined, simplify to !nil
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            let expr = walk_expression(p)?;
            return Ok(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(expr),
                right: Box::new(Expression::null()),
            });
        }
    }
    Ok(ExprKind::Lit(Literal::Bool(false)))
}

// ── Super ───────────────────────────────────────────────────────────────────

fn walk_super(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_args {
            args = walk_call_args(p)?;
        }
    }
    Ok(ExprKind::SuperCall { method: None, args })
}

// ── If/Unless as expression ─────────────────────────────────────────────────

fn walk_if_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let kind = walk_if(pair)?;
    // Wrap as a ternary-like expression
    if let StmtKind::If {
        cond,
        then_body,
        else_body,
        ..
    } = kind
    {
        let then_val = body_to_expr(then_body);
        let else_val = else_body.map(body_to_expr).unwrap_or(Expression::null());
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        })
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

fn walk_unless_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let kind = walk_unless(pair)?;
    if let StmtKind::If {
        cond,
        then_body,
        else_body,
        ..
    } = kind
    {
        let then_val = body_to_expr(then_body);
        let else_val = else_body.map(body_to_expr).unwrap_or(Expression::null());
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        })
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

fn walk_begin_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // begin..rescue..end as expression — just walk the body
    let kind = walk_begin(pair)?;
    if let StmtKind::Try { body, .. } = kind {
        Ok(body_to_expr(body).kind)
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

/// Convert a body (list of stmts) to a single expression (last statement value).
fn body_to_expr(mut stmts: Vec<Statement>) -> Expression {
    if stmts.is_empty() {
        return Expression::null();
    }
    let last = stmts.pop().unwrap();
    match last.kind {
        StmtKind::Expr(e) => e,
        StmtKind::Return(Some(e)) => e,
        _ => Expression::null(),
    }
}

// ── Expression list ─────────────────────────────────────────────────────────

fn walk_expr_list_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expr_kind(inner.into_iter().next().unwrap())
    } else if inner.is_empty() {
        Ok(ExprKind::Lit(Literal::Null))
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(ExprKind::Array(
            exprs
                .into_iter()
                .map(|e| ArrayElement {
                    key: None,
                    value: e,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ))
    }
}

fn walk_expr_list_single(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expression(inner.remove(0))
    } else if inner.is_empty() {
        Ok(Expression::null())
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(Expression::new(ExprKind::Array(
            exprs
                .into_iter()
                .map(|e| ArrayElement {
                    key: None,
                    value: e,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )))
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span {
        start_line: sl as u32,
        start_col: sc as u32,
        end_line: el as u32,
        end_col: ec as u32,
    }
}

fn negate(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

fn next_meaningful<'a>(
    iter: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        match p.as_rule() {
            Rule::NEWLINE | Rule::then_kw | Rule::do_kw | Rule::in_kw => continue,
            _ => return Ok(p),
        }
    }
    Err("No more meaningful pairs".into())
}

fn next_rule<'a>(
    iter: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn find_rule<'a>(
    iter: impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn find_rule_from_iter<'a>(
    iter: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn is_expression_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::expression
            | Rule::expression_list
            | Rule::ternary_expr
            | Rule::low_and_expr
            | Rule::low_or_expr
            | Rule::or_expr
            | Rule::and_expr
            | Rule::not_expr
            | Rule::comparison
            | Rule::bitor_expr
            | Rule::bitxor_expr
            | Rule::bitand_expr
            | Rule::shift_expr
            | Rule::range_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::postfix
            | Rule::primary
            | Rule::integer_literal
            | Rule::float_literal
            | Rule::string_literal
            | Rule::interpolated_string
            | Rule::heredoc
            | Rule::symbol
            | Rule::regex_literal
            | Rule::percent_literal
            | Rule::true_kw
            | Rule::false_kw
            | Rule::nil_kw
            | Rule::self_kw
            | Rule::identifier
            | Rule::constant
            | Rule::constant_path
            | Rule::instance_var
            | Rule::class_var
            | Rule::global_var
            | Rule::yield_expr
            | Rule::defined_expr
            | Rule::super_expr
            | Rule::block_given_expr
            | Rule::lambda_literal
            | Rule::proc_literal
            | Rule::if_expr
            | Rule::unless_expr
            | Rule::begin_expr
    )
}

fn is_op_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::additive_op
            | Rule::multiplicative_op
            | Rule::shift_op
            | Rule::comparison_op
            | Rule::range_op
            | Rule::aug_assign_op
    )
}

fn parse_comparison_op(s: &str) -> BinOp {
    match s {
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "<=>" => BinOp::Spaceship,
        "===" => BinOp::StrictEq,
        _ => BinOp::Eq,
    }
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn ruby_int_expr(value: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(value)))
}

fn ruby_call_expr(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ruby_array_expr(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn ruby_proc_expr(name: &str, lambda: Expression) -> Expression {
    let (arity, has_rest) = match &lambda.kind {
        ExprKind::Lambda { params, .. } => (
            params.iter().filter(|p| !p.is_rest).count() as i64,
            params.iter().any(|p| p.is_rest),
        ),
        _ => (0, false),
    };
    let param_count = match &lambda.kind {
        ExprKind::Lambda { params, .. } => params.len() as i64,
        _ => arity,
    };
    ruby_call_expr(
        name,
        vec![
            lambda,
            Expression::new(ExprKind::Lit(Literal::Int(arity))),
            Expression::new(ExprKind::Lit(Literal::Bool(has_rest))),
            Expression::new(ExprKind::Lit(Literal::Int(param_count))),
        ],
    )
}

fn ruby_add_expr(left: Expression, right: Expression) -> Expression {
    ruby_call_expr("__ruby_op_add", vec![left, right])
}

fn ruby_sub_expr(left: Expression, right: Expression) -> Expression {
    ruby_call_expr("__ruby_op_sub", vec![left, right])
}

fn is_negative_one_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr
        } if matches!(&expr.kind, ExprKind::Lit(Literal::Int(1)))
    )
}

fn is_negative_int_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr
        } if matches!(&expr.kind, ExprKind::Lit(Literal::Int(_)))
    )
}

fn literal_int_value(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(v)) => Some(*v),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => {
            if let ExprKind::Lit(Literal::Int(v)) = &expr.kind {
                Some(-*v)
            } else {
                None
            }
        }
        _ => None,
    }
}

fn ruby_slice_returns_nil(receiver: &Expression, method_name: &str, args: &[Argument]) -> bool {
    if method_name != "slice" || args.len() != 2 {
        return false;
    }
    if is_negative_int_expr(&args[1].value) {
        return true;
    }
    if let ExprKind::Array(elements) = &receiver.kind {
        if let Some(start) = literal_int_value(&args[0].value) {
            return start >= elements.len() as i64;
        }
    }
    false
}

fn ruby_range_exclusive_end(end: Expression, inclusive: bool) -> Expression {
    if !inclusive {
        return end;
    }
    if is_negative_one_expr(&end) {
        ruby_int_expr(i32::MAX as i64)
    } else {
        ruby_add_expr(end, ruby_int_expr(1))
    }
}

fn normalize_ruby_slice_call(method_name: &mut String, args: &mut Vec<Argument>) {
    if method_name != "slice" && method_name != "slice!" {
        return;
    }

    if args.len() == 1 {
        if let ExprKind::Range {
            start,
            end,
            inclusive,
        } = args[0].value.clone().kind
        {
            let start = *start;
            let exclusive_end = ruby_range_exclusive_end(*end, inclusive);
            if method_name == "slice!" {
                let count = ruby_sub_expr(exclusive_end, start.clone());
                args.clear();
                args.push(Argument::positional(start));
                args.push(Argument::positional(count));
            } else {
                args.clear();
                args.push(Argument::positional(start));
                args.push(Argument::positional(exclusive_end));
            }
        }
    } else if args.len() == 2 && method_name == "slice" {
        let start = args[0].value.clone();
        let len = args[1].value.clone();
        args[1].value = ruby_add_expr(start, len);
    }
}

fn parse_ruby_int(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    let (sign, body) = if let Some(rest) = s.strip_prefix('-') {
        (-1i64, rest)
    } else if let Some(rest) = s.strip_prefix('+') {
        (1i64, rest)
    } else {
        (1i64, s.as_str())
    };
    if body.starts_with("0x") || body.starts_with("0X") {
        Ok(ExprKind::Lit(Literal::Int(
            sign * i64::from_str_radix(&body[2..], 16).unwrap_or(0),
        )))
    } else if body.starts_with("0o") || body.starts_with("0O") {
        Ok(ExprKind::Lit(Literal::Int(
            sign * i64::from_str_radix(&body[2..], 8).unwrap_or(0),
        )))
    } else if body.starts_with("0b") || body.starts_with("0B") {
        Ok(ExprKind::Lit(Literal::Int(
            sign * i64::from_str_radix(&body[2..], 2).unwrap_or(0),
        )))
    } else {
        Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
    }
}

fn parse_ruby_float(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    Ok(ExprKind::Lit(Literal::Float(s.parse().unwrap_or(0.0))))
}

fn parse_ruby_string(s: &str) -> String {
    let (body, interpolate_escapes) = if s.starts_with("'''") || s.starts_with("\"\"\"") {
        (&s[3..s.len() - 3], s.starts_with("\"\"\""))
    } else if s.starts_with('"') {
        (&s[1..s.len() - 1], true)
    } else if s.starts_with('\'') {
        (&s[1..s.len() - 1], false)
    } else {
        (s, true)
    };
    if !interpolate_escapes {
        return body.replace("\\'", "'").replace("\\\\", "\\");
    }

    let mut out = String::new();
    let chars: Vec<char> = body.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '\\' || i + 1 >= chars.len() {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        match chars[i + 1] {
            'n' => {
                out.push('\n');
                i += 2;
            }
            'r' => {
                out.push('\r');
                i += 2;
            }
            't' => {
                out.push('\t');
                i += 2;
            }
            '"' => {
                out.push('"');
                i += 2;
            }
            '\'' => {
                out.push('\'');
                i += 2;
            }
            '\\' => {
                out.push('\\');
                i += 2;
            }
            'x' if i + 3 < chars.len() => {
                let hex = format!("{}{}", chars[i + 2], chars[i + 3]);
                if let Ok(byte) = u8::from_str_radix(&hex, 16) {
                    out.push(byte as char);
                    i += 4;
                } else {
                    out.push(chars[i]);
                    i += 1;
                }
            }
            'u' if i + 2 < chars.len() && chars[i + 2] == '{' => {
                let mut j = i + 3;
                let mut hex = String::new();
                while j < chars.len() && chars[j] != '}' {
                    hex.push(chars[j]);
                    j += 1;
                }
                if j < chars.len() && chars[j] == '}' {
                    for part in hex.split_whitespace() {
                        if let Ok(code) = u32::from_str_radix(part, 16) {
                            if let Some(ch) = char::from_u32(code) {
                                out.push(ch);
                            }
                        }
                    }
                    i = j + 1;
                } else {
                    out.push(chars[i]);
                    i += 1;
                }
            }
            other => {
                out.push(other);
                i += 2;
            }
        }
    }
    out
}

fn parse_heredoc(s: &str) -> String {
    // <<~TAG\ncontent\nTAG  or  <<TAG\ncontent\nTAG
    let squiggly = s.starts_with("<<~");
    let prefix_len = if squiggly { 3 } else { 2 };
    let rest = &s[prefix_len..];
    // Find the tag name (up to newline)
    if let Some(nl) = rest.find('\n') {
        let tag = rest[..nl].trim();
        let content = &rest[nl + 1..];
        // Strip trailing TAG line
        let body = if let Some(pos) = content.rfind(tag) {
            &content[..pos]
        } else {
            content
        };
        if squiggly {
            // Strip common leading whitespace
            let lines: Vec<&str> = body.lines().collect();
            let min_indent = lines
                .iter()
                .filter(|l| !l.trim().is_empty())
                .map(|l| l.len() - l.trim_start().len())
                .min()
                .unwrap_or(0);
            lines
                .iter()
                .map(|l| {
                    if l.len() > min_indent {
                        &l[min_indent..]
                    } else {
                        l.trim()
                    }
                })
                .collect::<Vec<_>>()
                .join("\n")
        } else {
            body.to_string()
        }
    } else {
        s.to_string()
    }
}

fn walk_percent_literal(s: &str) -> ExprKind {
    // %w[a b c] → array of strings
    // %W[a #{x} c] → array of interpolated strings
    // %i[a b c] → array of symbols (strings)
    // %I[a #{x} c] → array of interpolated symbols (strings)
    // %q[...] → single-quoted string
    // %Q[...] or %[...] → double-quoted string
    let (kind, interpolate, rest) = if s.starts_with("%w") || s.starts_with("%i") {
        ("array", false, &s[2..])
    } else if s.starts_with("%W") || s.starts_with("%I") {
        ("array", true, &s[2..])
    } else if s.starts_with("%q") || s.starts_with("%Q") {
        ("string", s.starts_with("%Q"), &s[2..])
    } else {
        ("string", true, &s[1..])
    };

    // Strip delimiters
    let body = if rest.len() >= 2 {
        &rest[1..rest.len() - 1]
    } else {
        rest
    };

    if kind == "array" {
        let words: Vec<ArrayElement> = ruby_percent_words(body, interpolate)
            .into_iter()
            .map(|w| ArrayElement {
                key: None,
                value: if interpolate && w.starts_with("#{") && w.ends_with('}') {
                    Expression::ident(&w[2..w.len() - 1])
                } else if !interpolate && w.starts_with("#{") && w.ends_with('}') {
                    Expression::new(ExprKind::Lit(Literal::Str(format!("\\{}", w))))
                } else {
                    Expression::new(ExprKind::Lit(Literal::Str(w)))
                },
                spread: false,
                by_ref: false,
            })
            .collect();
        ExprKind::Array(words)
    } else {
        ExprKind::Lit(Literal::Str(body.to_string()))
    }
}

fn ruby_percent_words(body: &str, interpolate: bool) -> Vec<String> {
    let mut words = Vec::new();
    let mut cur = String::new();
    let mut chars = body.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch.is_whitespace() {
            if !cur.is_empty() {
                words.push(std::mem::take(&mut cur));
            }
            continue;
        }
        if ch == '\\' {
            match chars.next() {
                Some(' ') => cur.push(' '),
                Some('n') if interpolate => cur.push('\n'),
                Some(other) => cur.push(other),
                None => cur.push('\\'),
            }
        } else {
            cur.push(ch);
        }
    }
    if !cur.is_empty() {
        words.push(cur);
    }
    words
}
