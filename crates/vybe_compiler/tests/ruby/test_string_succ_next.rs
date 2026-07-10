use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_succ_basic, "puts 'a'.succ", "b");
ruby_test!(test_string_succ_wrap, "puts 'z'.succ", "aa");
ruby_test!(test_string_succ_number, "puts '9'.succ", "10");
ruby_test!(test_string_succ_alphanumeric, "puts 'a9'.succ", "b0");
ruby_test!(test_string_succ_bang, "s = 'a'; s.succ!; puts s", "b");
ruby_test!(test_string_next_basic, "puts 'a'.next", "b");
ruby_test!(test_string_next_bang, "s = 'a'; s.next!; puts s", "b");
ruby_test!(test_string_upto, "acc = []; 'a'.upto('c') { |s| acc << s }; puts acc.join('-')", "a-b-c");
ruby_test!(test_string_upto_exclusive, "acc = []; 'a'.upto('c', true) { |s| acc << s }; puts acc.join('-')", "a-b");
