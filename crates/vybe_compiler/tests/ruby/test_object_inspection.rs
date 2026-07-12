macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_object_inspect_basic,
    "class A; end; puts A.new.inspect.start_with?('#<A:')",
    "true"
);
ruby_test!(
    test_object_inspect_override,
    "class A; def inspect; 'A-inspect'; end; end; puts A.new.inspect",
    "A-inspect"
);
ruby_test!(
    test_object_to_s_basic,
    "class A; end; puts A.new.to_s.start_with?('#<A:')",
    "true"
);
ruby_test!(
    test_object_to_s_override,
    "class A; def to_s; 'A-to_s'; end; end; puts A.new.to_s",
    "A-to_s"
);
ruby_test!(
    test_object_display,
    "class A; def to_s; 'A'; end; end; A.new.display",
    "A"
); // display prints to_s
