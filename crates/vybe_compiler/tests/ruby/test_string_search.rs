macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_string_search_index, "puts 'hello'.index('l')", "2");
ruby_test!(
    test_string_search_index_offset,
    "puts 'hello'.index('l', 3)",
    "3"
);
ruby_test!(
    test_string_search_index_not_found,
    "puts 'hello'.index('x').nil?",
    "true"
);
ruby_test!(test_string_search_rindex, "puts 'hello'.rindex('l')", "3");
ruby_test!(
    test_string_search_rindex_offset,
    "puts 'hello'.rindex('l', 2)",
    "2"
);
ruby_test!(
    test_string_search_match,
    "puts 'hello'.match('ell').to_a.join('-')",
    "ell"
);
ruby_test!(
    test_string_search_match_question,
    "puts 'hello'.match?('ell')",
    "true"
);
ruby_test!(
    test_string_search_scan,
    "puts 'hello'.scan('l').join('-')",
    "l-l"
);
ruby_test!(
    test_string_search_scan_block,
    "acc = []; 'hello'.scan(/./) { |c| acc << c }; puts acc.join('-')",
    "h-e-l-l-o"
);
