# vybe-test: ruby/modules_mixins/mixin_default_implementation_overridable
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Printable
  def to_print; self.class.to_s + ': default'; end
end
class Report
  include Printable
  def to_print; 'Report: custom'; end
end
class Invoice
  include Printable
end
puts Report.new.to_print
puts Invoice.new.to_print
