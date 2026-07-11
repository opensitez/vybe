
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_initialize_basic, "class A; def initialize; @x = 1; end; def x; @x; end; end; puts A.new.x", "1");
ruby_test!(test_initialize_args, "class A; def initialize(x); @x = x; end; def x; @x; end; end; puts A.new(2).x", "2");
ruby_test!(test_initialize_super, "class A; def initialize(x); @x = x; end; end; class B < A; def initialize(x, y); super(x); @y = y; end; def xy; \"#{@x}-#{@y}\"; end; end; puts B.new(1, 2).xy", "1-2");
ruby_test!(test_initialize_private_by_default, "class A; def initialize; end; end; puts A.new.private_methods.include?(:initialize)", "true");
ruby_test!(test_initialize_dup, "class A; attr_accessor :x; def initialize_dup(other); super; @x = other.x * 2; end; end; a = A.new; a.x = 2; b = a.dup; puts b.x", "4");
ruby_test!(test_initialize_clone, "class A; attr_accessor :x; def initialize_clone(other); super; @x = other.x * 3; end; end; a = A.new; a.x = 2; b = a.clone; puts b.x", "6");
