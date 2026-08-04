# vybe-test: ruby/functions/proc_new
# origin: languages/ruby/tests/ruby/test_functions.rs
# vybe-test-mode: compile

p = Proc.new { |x| x + 1 }
