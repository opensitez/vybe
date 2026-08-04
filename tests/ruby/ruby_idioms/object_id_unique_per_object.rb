# vybe-test: ruby/ruby_idioms/object_id_unique_per_object
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

a = 'hello'
b = 'hello'
result = a.object_id == b.object_id
