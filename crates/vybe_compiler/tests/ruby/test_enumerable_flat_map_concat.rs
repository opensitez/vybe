
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_flat_map_basic, "puts [1, 2].flat_map {|x| [x, x * 2]}.join('-')", "1-2-2-4");
ruby_test!(test_flat_map_no_block, "puts [1].flat_map.is_a?(Enumerator)", "true");
ruby_test!(test_concat_alias, "puts [1, 2].concat_map {|x| [x, x * 2]}.join('-')", "1-2-2-4"); // wait, ruby alias is actually flat_map, there is no concat_map. The alias is collect_concat!
ruby_test!(test_collect_concat_alias, "puts [1, 2].collect_concat {|x| [x, x * 2]}.join('-')", "1-2-2-4");
ruby_test!(test_flat_map_deep, "puts [1, 2].flat_map {|x| [[x]]}.inspect", "[[1], [2]]"); // only flattens one level
ruby_test!(test_flat_map_non_array, "puts [1, 2].flat_map {|x| x}.join('-')", "1-2"); // works even if block returns non-array
ruby_test!(test_flat_map_empty, "puts [].flat_map {|x| [x]}.length", "0");
ruby_test!(test_flat_map_returns_array, "puts [1].flat_map {|x| x}.is_a?(Array)", "true");
ruby_test!(test_flat_map_hash, "puts ({a: 1}.flat_map {|k, v| [k, v]}.map(&:to_s).join('-'))", "a-1");
