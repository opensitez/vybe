use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_while_basic, "i = 0; while i < 3; i += 1; end; puts i", "3");
ruby_test!(test_while_modifier, "i = 0; i += 1 while i < 3; puts i", "3");
ruby_test!(test_until_basic, "i = 0; until i == 3; i += 1; end; puts i", "3");
ruby_test!(test_until_modifier, "i = 0; i += 1 until i == 3; puts i", "3");
ruby_test!(test_loop_basic, "i = 0; loop do i += 1; break if i == 3; end; puts i", "3");
ruby_test!(test_while_false, "i = 0; while false; i += 1; end; puts i", "0");
ruby_test!(test_begin_while, "i = 0; begin; i += 1; end while false; puts i", "1"); // runs at least once
