
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_names_basic, "puts /(?<a>.)/.names.join('-')", "a");
ruby_test!(test_regex_names_multiple, "puts /(?<a>.)(?<b>.)/.names.join('-')", "a-b");
ruby_test!(test_regex_names_none, "puts /(.)/.names.length", "0");
ruby_test!(test_regex_named_captures_basic, "puts /(?<a>.)(?<b>.)/.named_captures.map{|k, v| \"#{k}:#{v.join(',')}\"}.join('-')", "a:1-b:2");
ruby_test!(test_regex_named_captures_multiple_same_name, "puts /(?<a>.)|(?<a>.)/.named_captures.map{|k, v| \"#{k}:#{v.join(',')}\"}.join('-')", "a:1,2");
