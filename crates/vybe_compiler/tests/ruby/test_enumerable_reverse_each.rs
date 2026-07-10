use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_reverse_each_basic, "acc = []; [1, 2, 3].reverse_each {|x| acc << x}; puts acc.join('-')", "3-2-1");
ruby_test!(test_reverse_each_no_block, "puts [1, 2, 3].reverse_each.is_a?(Enumerator)", "true");
ruby_test!(test_reverse_each_no_block_array, "puts [1, 2, 3].reverse_each.to_a.join('-')", "3-2-1");
ruby_test!(test_reverse_each_empty, "acc = []; [].reverse_each {|x| acc << x}; puts acc.length", "0");
ruby_test!(test_reverse_each_returns_self, "a = [1]; puts a.reverse_each {|x| x}.object_id == a.object_id", "true");
ruby_test!(test_reverse_each_range, "acc = []; (1..3).reverse_each {|x| acc << x}; puts acc.join('-')", "3-2-1");
ruby_test!(test_reverse_each_hash, "acc = []; {a: 1, b: 2}.reverse_each {|k, v| acc << k.to_s}; puts acc.join('-')", "b-a");
ruby_test!(test_reverse_each_string_chars, "acc = []; 'abc'.each_char.reverse_each {|c| acc << c}; puts acc.join('-')", "c-b-a"); // using enumerator reverse_each
