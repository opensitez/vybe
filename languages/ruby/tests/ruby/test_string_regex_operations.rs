macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_regex_match_operator,
    "puts 'hello' =~ /ll/ ? 'yes' : 'no'",
    "yes"
);
ruby_test!(
    test_regex_match_operator_fail,
    "puts 'hello' =~ /zz/ ? 'yes' : 'no'",
    "no"
);
ruby_test!(test_regex_match_index, "puts 'hello' =~ /ll/", "2");
ruby_test!(test_regex_match_method, "puts 'hello'.match(/ll/)[0]", "ll");
ruby_test!(
    test_regex_match_method_fail,
    "puts 'hello'.match(/zz/).nil?",
    "true"
);
ruby_test!(
    test_regex_match_predicate,
    "puts 'hello'.match?(/ll/)",
    "true"
);
ruby_test!(
    test_regex_match_predicate_fail,
    "puts 'hello'.match?(/zz/)",
    "false"
);
ruby_test!(
    test_regex_scan,
    "puts 'abacada'.scan(/a./).join('-')",
    "ab-ac-ad"
);
ruby_test!(
    test_regex_scan_groups,
    "puts 'abacada'.scan(/(a)(.)/).map{|x| x.join}.join('-')",
    "ab-ac-ad"
);
ruby_test!(
    test_regex_split,
    "puts 'a-b-c'.split(/-/).join(',')",
    "a,b,c"
);
ruby_test!(test_regex_gsub, "puts 'abacada'.gsub(/a./, 'X')", "XXa");
ruby_test!(test_regex_sub, "puts 'abacada'.sub(/a./, 'X')", "Xacada");
