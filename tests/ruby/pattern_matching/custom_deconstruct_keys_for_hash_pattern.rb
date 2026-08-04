# vybe-test: ruby/pattern_matching/custom_deconstruct_keys_for_hash_pattern
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


class Config
  def initialize(h, p); @host = h; @port = p; end
  def deconstruct_keys(keys); { host: @host, port: @port }; end
end
case Config.new("localhost", 8080)
in { host: String => h, port: Integer => p }
  puts h.to_s + ':' + p.to_s
end
