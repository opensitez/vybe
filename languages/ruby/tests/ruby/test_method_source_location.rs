macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_method_source_location_basic,
    "class A; def foo; end; end; puts A.new.method(:foo).source_location[0].end_with?('.rb') || A.new.method(:foo).source_location[0] == '-e'",
    "true"
); // In run_ruby_one, it might be '-e'
ruby_test!(
    test_method_source_location_line,
    "class A\n def foo; end\n end; puts A.new.method(:foo).source_location[1]",
    "2"
);
ruby_test!(
    test_unbound_method_source_location,
    "class A\n def foo; end\n end; puts A.instance_method(:foo).source_location[1]",
    "2"
);
ruby_test!(
    test_method_source_location_native,
    "puts [].method(:push).source_location.nil?",
    "true"
); // Native methods return nil
