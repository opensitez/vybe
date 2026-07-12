macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_bsearch_index_exact,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| 5 <=> x}",
    "2"
);
ruby_test!(
    test_bsearch_index_boolean_first,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| x >= 4}",
    "2"
); // index of 5
ruby_test!(
    test_bsearch_index_boolean_all_false,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| x >= 10}.nil?",
    "true"
);
ruby_test!(
    test_bsearch_index_boolean_all_true,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| x >= 0}",
    "0"
);
ruby_test!(
    test_bsearch_index_spaceship_found,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| 7 <=> x}",
    "3"
);
ruby_test!(
    test_bsearch_index_spaceship_not_found,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| 6 <=> x}.nil?",
    "true"
);
ruby_test!(
    test_bsearch_index_spaceship_first,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| 1 <=> x}",
    "0"
);
ruby_test!(
    test_bsearch_index_spaceship_last,
    "puts [1, 3, 5, 7, 9].bsearch_index {|x| 9 <=> x}",
    "4"
);
ruby_test!(
    test_bsearch_index_empty,
    "puts [].bsearch_index {|x| x >= 1}.nil?",
    "true"
);
ruby_test!(
    test_bsearch_index_duplicate_boolean,
    "idx = [1, 5, 5, 5, 9].bsearch_index {|x| x >= 4}; puts idx > 0 && idx < 4",
    "true"
); // might return 1, 2 or 3 depending on impl, usually the exact middle if it hits
ruby_test!(
    test_bsearch_index_string,
    "puts ['a', 'c', 'e'].bsearch_index {|x| x >= 'b'}",
    "1"
);
ruby_test!(
    test_bsearch_index_out_of_range,
    "puts [1, 3].bsearch_index {|x| 5 <=> x}.nil?",
    "true"
);
