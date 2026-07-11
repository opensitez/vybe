
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_for_basic, "acc = []; for i in [1, 2, 3]; acc << i; end; puts acc.join('-')", "1-2-3");
ruby_test!(test_for_range, "acc = []; for i in 1..3; acc << i; end; puts acc.join('-')", "1-2-3");
ruby_test!(test_for_hash, "acc = []; for k, v in {a: 1, b: 2}; acc << \"#{k}#{v}\"; end; puts acc.join('-')", "a1-b2");
ruby_test!(test_for_variable_leaks, "for i in [1]; end; puts i", "1"); // loop variable leaks to outer scope
ruby_test!(test_for_multiple_assignment, "acc = []; for a, b in [[1, 2], [3, 4]]; acc << \"#{a}-#{b}\"; end; puts acc.join('|')", "1-2|3-4");
