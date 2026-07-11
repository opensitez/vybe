
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_cmp_equal, "puts (Time.at(100) <=> Time.at(100))", "0");
ruby_test!(test_time_cmp_less, "puts (Time.at(100) <=> Time.at(200))", "-1");
ruby_test!(test_time_cmp_greater, "puts (Time.at(200) <=> Time.at(100))", "1");
ruby_test!(test_time_cmp_type_mismatch, "puts (Time.at(100) <=> 100).nil?", "true");
ruby_test!(test_time_eq_basic, "puts (Time.at(100) == Time.at(100))", "true");
ruby_test!(test_time_eq_false, "puts (Time.at(100) == Time.at(200))", "false");
ruby_test!(test_time_eq_type_mismatch, "puts (Time.at(100) == 100)", "false");
ruby_test!(test_time_eql_basic, "puts Time.at(100).eql?(Time.at(100))", "true");
ruby_test!(test_time_hash_equal, "puts Time.at(100).hash == Time.at(100).hash", "true");
ruby_test!(test_time_hash_diff, "puts Time.at(100).hash == Time.at(200).hash", "false");
ruby_test!(test_time_between, "puts Time.at(150).between?(Time.at(100), Time.at(200))", "true");
