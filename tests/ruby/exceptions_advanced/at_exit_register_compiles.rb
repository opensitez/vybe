# vybe-test: ruby/exceptions_advanced/at_exit_register_compiles
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

at_exit { puts 'cleanup on exit' }
