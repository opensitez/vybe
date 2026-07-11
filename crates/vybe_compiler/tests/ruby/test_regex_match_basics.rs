
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_match_basic, "puts /a/ =~ 'cat'", "1");
ruby_test!(test_regex_match_missing, "puts (/a/ =~ 'dog').nil?", "true");
ruby_test!(test_regex_match_string, "puts 'cat' =~ /a/", "1");
ruby_test!(test_regex_match_operator, "puts /a/.match?('cat')", "true"); // ruby 2.4+ match?
ruby_test!(test_regex_match_operator_false, "puts /a/.match?('dog')", "false");
ruby_test!(test_regex_match_operator_pos, "puts /a/.match?('cat', 2)", "false"); // start at pos 2
ruby_test!(test_regex_match_method, "puts /a/.match('cat')[0]", "a");
ruby_test!(test_regex_match_method_missing, "puts /a/.match('dog').nil?", "true");
ruby_test!(test_regex_eqq_basic, "puts /a/ === 'cat'", "true");
ruby_test!(test_regex_eqq_false, "puts /a/ === 'dog'", "false");
