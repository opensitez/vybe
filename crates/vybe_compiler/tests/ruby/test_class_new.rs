
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_class_new_basic, "c = Class.new; puts c.class.name", "Class");
ruby_test!(test_class_new_superclass, "c = Class.new(Array); puts c.new.is_a?(Array)", "true");
ruby_test!(test_class_new_block, "c = Class.new { def foo; 42; end }; puts c.new.foo", "42");
ruby_test!(test_class_allocate, "class C; def initialize; @x = 1; end; attr_reader :x; end; puts C.allocate.x.nil?", "true");
ruby_test!(test_class_superclass, "puts Array.superclass == Object", "true");
ruby_test!(test_class_superclass_basic_object, "puts BasicObject.superclass.nil?", "true");
ruby_test!(test_class_subclasses, "class C; end; class D < C; end; puts C.subclasses.include?(D).to_s", "true");
ruby_test!(test_class_attached_object, "class C; class << self; puts self.attached_object == C; end; end", "true");
