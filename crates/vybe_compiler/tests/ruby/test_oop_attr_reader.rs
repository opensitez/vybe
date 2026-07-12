macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_attr_reader_basic,
    "class A; attr_reader :x; def initialize(x); @x = x; end; end; puts A.new(1).x",
    "1"
);
ruby_test!(
    test_attr_reader_multiple,
    "class A; attr_reader :x, :y; def initialize(x, y); @x = x; @y = y; end; end; a = A.new(1, 2); puts \"#{a.x}-#{a.y}\"",
    "1-2"
);
ruby_test!(
    test_attr_reader_string,
    "class A; attr_reader 'x'; def initialize(x); @x = x; end; end; puts A.new(1).x",
    "1"
);
ruby_test!(
    test_attr_reader_uninitialized,
    "class A; attr_reader :x; end; puts A.new.x.nil?",
    "true"
);
ruby_test!(
    test_attr_reader_missing_method_error,
    "class A; attr_reader :x; end; begin; A.new.x = 1; rescue NoMethodError; puts 'err'; end",
    "err"
);
