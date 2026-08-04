# vybe-test: ruby/modules_mixins/module_constant_access
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Config
  VERSION = "1.0.0"
  MAX_RETRIES = 3
end
puts Config::VERSION
puts Config::MAX_RETRIES
