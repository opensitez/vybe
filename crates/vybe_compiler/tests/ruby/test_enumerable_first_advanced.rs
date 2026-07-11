
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_first_basic, "puts [1, 2, 3].first", "1");
ruby_test!(test_first_empty, "puts [].first.nil?", "true");
ruby_test!(test_first_arg_one, "puts [1, 2, 3].first(1).join('-')", "1");
ruby_test!(test_first_arg_multiple, "puts [1, 2, 3].first(2).join('-')", "1-2");
ruby_test!(test_first_arg_zero, "puts [1, 2].first(0).length", "0");
ruby_test!(test_first_arg_all, "puts [1, 2].first(5).join('-')", "1-2");
ruby_test!(test_first_arg_negative_error, "begin; [1].first(-1); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_first_hash, "puts ({a: 1}.first.join('-'))", "a-1");
ruby_test!(test_first_hash_empty, "puts ({}).first.nil?", "true");
ruby_test!(test_first_hash_arg, "puts ({a: 1, b: 2}.first(2).map{|kv| kv.join(':')}.join('-'))", "a:1-b:2");
