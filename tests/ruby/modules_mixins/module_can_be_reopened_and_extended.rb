# vybe-test: ruby/modules_mixins/module_can_be_reopened_and_extended
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Formatter
  def shout; upcase + "!"; end
end
module Formatter
  def whisper; downcase + "..."; end
end
class String
  include Formatter
end
puts "hello".shout
puts "HELLO".whisper
