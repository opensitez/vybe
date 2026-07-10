use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_casecmp_equal, "puts 'abc'.casecmp('ABC')", "0");
ruby_test!(test_casecmp_less, "puts 'abc'.casecmp('ABD')", "-1");
ruby_test!(test_casecmp_greater, "puts 'abd'.casecmp('ABC')", "1");
ruby_test!(test_casecmp_different_lengths, "puts 'a'.casecmp('A ')", "-1");
ruby_test!(test_casecmp_type_error, "puts 'a'.casecmp(1).nil?", "true");
ruby_test!(test_casecmp_question_equal, "puts 'abc'.casecmp?('ABC')", "true");
ruby_test!(test_casecmp_question_not_equal, "puts 'abc'.casecmp?('ABD')", "false");
ruby_test!(test_casecmp_question_type_error, "puts 'a'.casecmp?(1).nil?", "true");
ruby_test!(test_casecmp_unicode, "puts 'é'.casecmp('É')", "0"); // Might depend on ruby version, usually works
ruby_test!(test_casecmp_question_unicode, "puts 'é'.casecmp?('É')", "true");
ruby_test!(test_cmp_equal, "puts 'abc' <=> 'abc'", "0");
ruby_test!(test_cmp_case_sensitive, "puts 'abc' <=> 'ABC'", "1"); // a > A in ascii
