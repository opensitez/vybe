
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_count_basic, "puts 'hello'.count('l')", "2");
ruby_test!(test_count_multiple_chars, "puts 'hello'.count('lo')", "3");
ruby_test!(test_count_range, "puts 'hello'.count('a-j')", "2"); // h, e
ruby_test!(test_count_negation, "puts 'hello'.count('^l')", "3"); // h, e, o
ruby_test!(test_count_negation_range, "puts 'hello'.count('^a-j')", "3"); // l, l, o
ruby_test!(test_count_intersection, "puts 'hello'.count('lo', 'o')", "1"); // intersection of ['l', 'o'] and ['o']
ruby_test!(test_count_intersection_negation, "puts 'hello'.count('lo', '^l')", "1"); // 'o'
ruby_test!(test_count_empty, "puts ''.count('a')", "0");
ruby_test!(test_count_not_found, "puts 'hello'.count('z')", "0");
ruby_test!(test_count_unicode, "puts 'éé'.count('é')", "2");
ruby_test!(test_count_escaped_dash, "puts 'a-b'.count('a\\\\-c')", "2"); // counting 'a', '-', 'c' -> 'a', '-' are in 'a-b'
