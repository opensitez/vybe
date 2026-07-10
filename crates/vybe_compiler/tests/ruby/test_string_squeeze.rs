use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_squeeze_basic, "puts 'yellow moon'.squeeze", "yelow mon");
ruby_test!(test_squeeze_specific_char, "puts 'yellow moon'.squeeze('o')", "yellow mon");
ruby_test!(test_squeeze_range, "puts 'putters shoot too'.squeeze('m-z')", "puters shot to");
ruby_test!(test_squeeze_negation, "puts 'putters shoot too'.squeeze('^o')", "puters shoot too"); // squeeze everything EXCEPT 'o'
ruby_test!(test_squeeze_multiple_args, "puts 'aabbaabb'.squeeze('a', 'b')", "abab");
ruby_test!(test_squeeze_bang_mutates, "s = 'yellow'; s.squeeze!; puts s", "yelow");
ruby_test!(test_squeeze_bang_returns_nil_if_no_change, "s = 'abc'; puts s.squeeze!.nil?", "true");
ruby_test!(test_squeeze_unicode, "puts 'ééé'.squeeze", "é");
ruby_test!(test_squeeze_empty, "puts ''.squeeze", "");
ruby_test!(test_squeeze_no_duplicates, "puts 'abc'.squeeze", "abc");
ruby_test!(test_squeeze_triple, "puts 'aaa'.squeeze", "a");
