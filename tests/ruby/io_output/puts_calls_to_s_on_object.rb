# vybe-test: ruby/io_output/puts_calls_to_s_on_object
# origin: languages/ruby/tests/ruby/test_io_output.rs
# vybe-test-mode: compile


class Widget
  def to_s; "Widget!"; end
end
puts Widget.new
