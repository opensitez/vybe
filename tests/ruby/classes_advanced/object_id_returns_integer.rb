# vybe-test: ruby/classes_advanced/object_id_returns_integer
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Widget
end
w = Widget.new
id = w.object_id
