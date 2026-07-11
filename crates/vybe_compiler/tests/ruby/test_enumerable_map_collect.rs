
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_map_basic, "puts [1, 2, 3].map {|x| x * 2}.join('-')", "2-4-6");
ruby_test!(test_map_no_block, "puts [1].map.is_a?(Enumerator)", "true");
ruby_test!(test_map_returns_array, "puts [1].map {|x| x}.is_a?(Array)", "true");
ruby_test!(test_collect_alias, "puts [1, 2, 3].collect {|x| x * 2}.join('-')", "2-4-6");
ruby_test!(test_map_empty, "puts [].map {|x| x}.length", "0");
ruby_test!(test_map_does_not_mutate, "a = [1]; a.map {|x| x * 2}; puts a[0]", "1");
ruby_test!(test_map_different_types, "puts [1, 'a'].map {|x| x.to_s * 2}.join('-')", "11-aa");
ruby_test!(test_map_nil_elements, "puts [nil, 1].map {|x| x.nil?}.join('-')", "true-false");
ruby_test!(test_map_hash, "puts ({a: 1, b: 2}.map {|k, v| \"#{k}:#{v}\"}.join('-'))", "a:1-b:2");
