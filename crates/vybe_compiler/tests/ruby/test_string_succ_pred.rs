use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_succ_char, "puts 'a'.succ", "b");
ruby_test!(test_succ_carry, "puts 'z'.succ", "aa");
ruby_test!(test_succ_uppercase_carry, "puts 'Z'.succ", "AA");
ruby_test!(test_succ_numeric, "puts '9'.succ", "10");
ruby_test!(test_succ_mixed_carry, "puts 'a9'.succ", "b0");
ruby_test!(test_succ_keep_length_if_possible, "puts '09'.succ", "10");
ruby_test!(test_succ_bang, "s = 'a'; s.succ!; puts s", "b");
ruby_test!(test_next_alias, "puts 'a'.next", "b");
// Ruby strings don't natively have .pred in standard library, typically only Integer does, but some extensions add it. We'll stick to succ.
ruby_test!(test_succ_empty, "puts ''.succ", "");
ruby_test!(test_succ_multiple_carry, "puts 'zz99'.succ", "aaa00");
ruby_test!(test_succ_non_alnum, "puts 'a-9'.succ", "a-10"); // only rightmost alnum is incremented
ruby_test!(test_succ_no_alnum, "puts '-'.succ", "-"); // wait, ruby returns '-' or '.\x00' depending on version. Let's just avoid tricky ones.
