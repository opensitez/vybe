
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_matchdata_captures_array, "m = /(a)(b)/.match('abc'); puts m.captures.join('-')", "a-b");
ruby_test!(test_matchdata_captures_bracket_zero, "m = /(a)(b)/.match('abc'); puts m[0]", "ab");
ruby_test!(test_matchdata_captures_bracket_index, "m = /(a)(b)/.match('abc'); puts m[1]", "a");
ruby_test!(test_matchdata_captures_bracket_out_of_bounds, "m = /(a)(b)/.match('abc'); puts m[3].nil?", "true");
ruby_test!(test_matchdata_captures_bracket_negative, "m = /(a)(b)/.match('abc'); puts m[-1]", "b");
ruby_test!(test_matchdata_captures_named, "m = /(?<a>a)(?<b>b)/.match('abc'); puts m['a']", "a");
ruby_test!(test_matchdata_captures_named_symbol, "m = /(?<a>a)(?<b>b)/.match('abc'); puts m[:a]", "a");
ruby_test!(test_matchdata_captures_named_captures_method, "m = /(?<a>a)(?<b>b)/.match('abc'); puts m.named_captures['a']", "a");
ruby_test!(test_matchdata_captures_names_method, "m = /(?<a>a)(?<b>b)/.match('abc'); puts m.names.join('-')", "a-b");
ruby_test!(test_matchdata_captures_to_a, "m = /(a)(b)/.match('abc'); puts m.to_a.join('-')", "ab-a-b");
