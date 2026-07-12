macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_const_missing_basic,
    "class A; def self.const_missing(c); \"missing #{c}\"; end; end; puts A::Foo",
    "missing Foo"
);
ruby_test!(
    test_const_missing_super,
    "class A; def self.const_missing(c); super; rescue NameError; 'err'; end; end; puts A::Foo",
    "err"
);
ruby_test!(
    test_const_missing_module,
    "module M; def self.const_missing(c); \"missing #{c}\"; end; end; puts M::Foo",
    "missing Foo"
);
ruby_test!(
    test_const_missing_global,
    "def Object.const_missing(c); \"missing #{c}\"; end; puts Foo",
    "missing Foo"
);
