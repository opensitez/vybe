use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_capture_basic, "m = /(.)(.)(.)/.match('abc'); puts \"#{m[1]}-#{m[2]}-#{m[3]}\"", "a-b-c");
ruby_test!(test_regex_capture_named, "m = /(?<a>.)(?<b>.)(?<c>.)/.match('abc'); puts \"#{m[:a]}-#{m[:b]}-#{m[:c]}\"", "a-b-c");
ruby_test!(test_regex_capture_global, "/(.)(.)(.)/ =~ 'abc'; puts \"#{$1}-#{$2}-#{$3}\"", "a-b-c");
ruby_test!(test_regex_capture_named_global, "/(?<a>.)(?<b>.)(?<c>.)/ =~ 'abc'; puts \"#{a}-#{b}-#{c}\"", "a-b-c"); // local variables assigned!
ruby_test!(test_regex_capture_pre_match, "/b/ =~ 'abc'; puts $`", "a");
ruby_test!(test_regex_capture_post_match, "/b/ =~ 'abc'; puts $'", "c");
ruby_test!(test_regex_capture_entire, "/b/ =~ 'abc'; puts $&", "b");
ruby_test!(test_regex_capture_last, "/(a)(b)(c)/ =~ 'abc'; puts $+", "c"); // $+ is last captured group
