macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_lineno_basic,
    "require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; puts t.lineno",
    "0"
);
ruby_test!(
    test_lineno_increment,
    "require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; t.gets; puts t.lineno",
    "1"
);
ruby_test!(
    test_lineno_set,
    "require 'tempfile'; t = Tempfile.new('ln'); t.lineno = 5; puts t.lineno",
    "5"
);
ruby_test!(
    test_lineno_increment_after_set,
    "require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; t.lineno = 5; t.gets; puts t.lineno",
    "6"
);
ruby_test!(
    test_dollar_dot,
    "require 'tempfile'; t = Tempfile.new('ln'); t.write(\"a\\nb\\nc\"); t.rewind; t.gets; puts $.",
    "1"
); // $. holds the lineno of the last read file
ruby_test!(test_dollar_dot_set, "$. = 10; puts $.", "10");
