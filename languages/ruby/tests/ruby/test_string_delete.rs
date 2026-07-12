macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_delete_basic, "puts 'hello'.delete('l')", "heo");
ruby_test!(
    test_delete_multiple_chars,
    "puts 'hello'.delete('lo')",
    "he"
);
ruby_test!(test_delete_range, "puts 'hello'.delete('a-j')", "ll"); // deletes h, e
ruby_test!(test_delete_negation, "puts 'hello'.delete('^l')", "ll"); // deletes everything EXCEPT l
ruby_test!(
    test_delete_negation_range,
    "puts 'hello'.delete('^a-j')",
    "he"
); // keeps h, e
ruby_test!(
    test_delete_intersection,
    "puts 'hello'.delete('lo', 'o')",
    "hell"
); // intersection of ['l', 'o'] and ['o'] is 'o'
ruby_test!(
    test_delete_intersection_negation,
    "puts 'hello'.delete('lo', '^l')",
    "hell"
); // deletes 'o'
ruby_test!(
    test_delete_bang_mutates,
    "s = 'hello'; s.delete!('l'); puts s",
    "heo"
);
ruby_test!(
    test_delete_bang_returns_nil,
    "s = 'hello'; puts s.delete!('z').nil?",
    "true"
);
ruby_test!(test_delete_empty, "puts ''.delete('a')", "");
ruby_test!(test_delete_not_found, "puts 'hello'.delete('z')", "hello");
ruby_test!(test_delete_unicode, "puts 'ééabc'.delete('é')", "abc");
