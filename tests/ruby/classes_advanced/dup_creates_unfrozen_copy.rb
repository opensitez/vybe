# vybe-test: ruby/classes_advanced/dup_creates_unfrozen_copy
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Config
  attr_accessor :setting
end
orig = Config.new
orig.freeze
copy = orig.dup
copy.setting = 'new'
