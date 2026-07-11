
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regexp_matching_match_method, "puts /l/.match('hello').to_a.join('-')", "l");
ruby_test!(test_regexp_matching_match_operator, "puts (/l/ =~ 'hello')", "2");
ruby_test!(test_regexp_matching_match_operator_reverse, "puts ('hello' =~ /l/)", "2");
ruby_test!(test_regexp_matching_match_not_found, "puts (/x/.match('hello').nil?)", "true");
ruby_test!(test_regexp_matching_match_operator_not_found, "puts (/x/ =~ 'hello').nil?", "true");
ruby_test!(test_regexp_matching_match_with_position, "puts /l/.match('hello', 3).to_a.join('-')", "l");
ruby_test!(test_regexp_matching_match_bang, "puts (/l/.match?('hello'))", "true");
ruby_test!(test_regexp_matching_match_bang_not_found, "puts (/x/.match?('hello'))", "false");
ruby_test!(test_regexp_matching_case_equality, "puts (/l/ === 'hello')", "true");
ruby_test!(test_regexp_matching_case_equality_not_found, "puts (/x/ === 'hello')", "false");
ruby_test!(test_regexp_matching_tilde, "puts (~/l/)", "nil"); // returns index if matching against $_, but $_ is nil here
