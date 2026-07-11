
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_gsub_basic, "puts 'hello'.gsub('l', 'r')", "herro");
ruby_test!(test_string_gsub_regex, "puts 'hello'.gsub(/[aeiou]/, '*')", "h*ll*");
ruby_test!(test_string_gsub_block, "puts 'hello'.gsub(/./) { |c| c.upcase }", "HELLO");
ruby_test!(test_string_gsub_hash, "puts 'hello'.gsub(/[eo]/, 'e' => 3, 'o' => 0)", "h3ll0");
ruby_test!(test_string_gsub_bang_basic, "s = 'hello'; s.gsub!('l', 'r'); puts s", "herro");
ruby_test!(test_string_gsub_bang_regex, "s = 'hello'; s.gsub!(/[aeiou]/, '*'); puts s", "h*ll*");
ruby_test!(test_string_gsub_bang_block, "s = 'hello'; s.gsub!(/./) { |c| c.upcase }; puts s", "HELLO");
ruby_test!(test_string_gsub_bang_hash, "s = 'hello'; s.gsub!(/[eo]/, 'e' => 3, 'o' => 0); puts s", "h3ll0");
ruby_test!(test_string_gsub_bang_no_match, "s = 'hello'; puts s.gsub!('z', 'r').nil?", "true");
ruby_test!(test_string_gsub_multiple_groups, "puts 'hello world'.gsub(/(h)(e)/, '\\\\2\\\\1')", "ehllo world");
ruby_test!(test_string_gsub_backslash, "puts 'hello'.gsub('l', '\\\\\\\\')", "he\\\\o");
ruby_test!(test_string_gsub_proc_replace, "puts 'hello'.gsub(/[aeiou]/, proc { |c| c.upcase })", "hEllO");
ruby_test!(test_string_gsub_named_captures, "puts 'hello'.gsub(/(?<vowel>[aeiou])/, '{\\\\k<vowel>}')", "h{e}ll{o}");
ruby_test!(test_string_sub_basic, "puts 'hello'.sub('l', 'r')", "herlo");
ruby_test!(test_string_sub_regex, "puts 'hello'.sub(/[aeiou]/, '*')", "h*llo");
ruby_test!(test_string_sub_bang, "s = 'hello'; s.sub!('l', 'r'); puts s", "herlo");
