# vybe-test: ruby/modules_mixins/mixin_instance_and_class_methods
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Persistable
  def self.included(base)
    base.extend(ClassMethods)
  end
  module ClassMethods
    def find(id); "Record #{id}"; end
  end
  def save; "saved"; end
end
class User
  include Persistable
end
puts User.find(1)
puts User.new.save
