# vybe-test: ruby/modules_mixins/ancestors_includes_modules
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Printable; end
module Serializable; end
class Document
  include Printable
  include Serializable
end
puts Document.ancestors.include?(Printable)
puts Document.ancestors.include?(Serializable)
