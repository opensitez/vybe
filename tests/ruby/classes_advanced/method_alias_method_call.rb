# vybe-test: ruby/classes_advanced/method_alias_method_call
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Printer
  def print_text
    'printing'
  end
  alias_method :display, :print_text
end
p = Printer.new
p.display
