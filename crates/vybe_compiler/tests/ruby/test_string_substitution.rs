
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_format_gsub, "puts 'hello'.gsub('l', 'r')", "herro");
ruby_test!(test_string_format_gsub_bang, "s = 'hello'; s.gsub!('l', 'r'); puts s", "herro");
ruby_test!(test_string_format_sub, "puts 'hello'.sub('l', 'r')", "herlo");
ruby_test!(test_string_format_sub_bang, "s = 'hello'; s.sub!('l', 'r'); puts s", "herlo");
ruby_test!(test_string_format_gsub_block, "puts 'hello'.gsub(/./) { |c| c.ord.to_s + '-' }", "104-101-108-108-111-");
ruby_test!(test_string_format_gsub_hash, "puts 'hello'.gsub(/[aeiou]/, 'e' => '3', 'o' => '0')", "h3ll0");
ruby_test!(test_string_format_sub_hash, "puts 'hello'.sub(/[aeiou]/, 'e' => '3', 'o' => '0')", "h3llo");
ruby_test!(test_string_format_sub_block, "puts 'hello'.sub(/./) { |c| c.upcase }", "Hello");
