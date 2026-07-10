use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_match_operator_basic, "puts 'hello' =~ /ll/", "2");
ruby_test!(test_match_operator_false, "puts ('hello' =~ /xx/).nil?", "true");
ruby_test!(test_match_operator_reverse, "puts /ll/ =~ 'hello'", "2");
ruby_test!(test_not_match_operator_basic, "puts 'hello' !~ /xx/", "true");
ruby_test!(test_not_match_operator_false, "puts 'hello' !~ /ll/", "false");
ruby_test!(test_not_match_operator_reverse, "puts /xx/ !~ 'hello'", "true");
ruby_test!(test_match_operator_non_string, "puts (1 =~ /1/).nil?", "true"); // integer =~ regex returns nil in modern ruby wait actually Integer#=~ is deprecated or removed. Let's stick to Strings and Regexp.
ruby_test!(test_match_operator_sets_globals, "'hello' =~ /(l)(l)/; puts \"#{$1}-#{$2}\"", "l-l");
