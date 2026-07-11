
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_scan_basic, "puts 'hello'.scan(/l/).join('-')", "l-l");
ruby_test!(test_string_scan_groups, "puts 'hello'.scan(/(.)(l)/).map{|g| g.join}.join('-')", "el-ll");
ruby_test!(test_string_scan_block, "acc = []; 'hello'.scan(/l/) { |m| acc << m }; puts acc.join('-')", "l-l");
ruby_test!(test_string_scan_string, "puts 'hello'.scan('l').join('-')", "l-l");
ruby_test!(test_string_scan_empty, "puts 'hello'.scan(/x/).length", "0");
ruby_test!(test_string_scan_overlapping, "puts 'aaaa'.scan(/aa/).join('-')", "aa-aa"); // scan doesn't overlap
ruby_test!(test_string_scan_multiple_groups, "acc = []; 'h1e2'.scan(/([a-z])([0-9])/) { |g1, g2| acc << \"#{g1}-#{g2}\" }; puts acc.join('|')", "h-1|e-2");
ruby_test!(test_string_scan_empty_string, "puts ''.scan(/./).length", "0");
ruby_test!(test_string_scan_entire_string, "puts 'hello'.scan(/.*/).join('-')", "hello-"); // empty match at end
ruby_test!(test_string_scan_named_captures, "acc = []; 'h1e2'.scan(/(?<letter>[a-z])(?<num>[0-9])/) { |m| acc << m.join('-') }; puts acc.join('|')", "h-1|e-2");
