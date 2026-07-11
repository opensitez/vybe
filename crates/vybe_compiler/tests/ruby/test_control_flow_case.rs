
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_case_basic, "x = 2; case x; when 1; puts 'a'; when 2; puts 'b'; else; puts 'c'; end", "b");
ruby_test!(test_case_multiple_matches, "x = 2; case x; when 1, 2; puts 'a'; else; puts 'c'; end", "a");
ruby_test!(test_case_else, "x = 3; case x; when 1; puts 'a'; when 2; puts 'b'; else; puts 'c'; end", "c");
ruby_test!(test_case_no_condition, "x = 2; case; when x == 1; puts 'a'; when x == 2; puts 'b'; else; puts 'c'; end", "b");
ruby_test!(test_case_class_match, "x = 'hello'; case x; when String; puts 's'; when Integer; puts 'i'; end", "s");
ruby_test!(test_case_regex_match, "x = 'hello'; case x; when /ll/; puts 'r'; else; puts 'no'; end", "r");
ruby_test!(test_case_range_match, "x = 5; case x; when 1..3; puts 'a'; when 4..6; puts 'b'; end", "b");
ruby_test!(test_case_then_keyword, "x = 1; case x; when 1 then puts 'a'; else puts 'b'; end", "a");
