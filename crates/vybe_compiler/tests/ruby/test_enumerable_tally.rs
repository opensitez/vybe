
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_tally_basic, "puts %w[a b a c b a].tally.map { |k, v| \"#{k}:#{v}\" }.sort.join('-')", "a:3-b:2-c:1");
ruby_test!(test_enumerable_tally_hash, "h = { 'a' => 1 }; puts %w[a b a].tally(h).map { |k, v| \"#{k}:#{v}\" }.sort.join('-')", "a:3-b:1");
ruby_test!(test_enumerable_uniq_basic, "puts [1, 2, 1, 3, 2].uniq.join('-')", "1-2-3");
ruby_test!(test_enumerable_uniq_block, "puts %w[a aa b bb c].uniq { |x| x.length }.join('-')", "a-aa");
ruby_test!(test_enumerable_sum_basic, "puts [1, 2, 3].sum", "6");
ruby_test!(test_enumerable_sum_init, "puts [1, 2, 3].sum(10)", "16");
ruby_test!(test_enumerable_sum_block, "puts [1, 2, 3].sum { |x| x * 2 }", "12");
