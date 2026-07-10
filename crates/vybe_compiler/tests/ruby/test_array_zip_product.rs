use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_zip_basic, "puts [1, 2].zip([3, 4]).map { |a| a.join('') }.join('-')", "13-24");
ruby_test!(test_array_zip_uneven, "puts [1, 2].zip([3]).map { |a| a.map(&:to_s).join('') }.join('-')", "13-2");
ruby_test!(test_array_zip_multiple, "puts [1, 2].zip([3, 4], [5, 6]).map { |a| a.join('') }.join('-')", "135-246");
ruby_test!(test_array_zip_block, "acc = []; [1, 2].zip([3, 4]) { |a, b| acc << a + b }; puts acc.join('-')", "4-6");
ruby_test!(test_array_product_basic, "puts [1, 2].product([3, 4]).map { |a| a.join('') }.join('-')", "13-14-23-24");
ruby_test!(test_array_product_multiple, "puts [1, 2].product([3], [4]).map { |a| a.join('') }.join('-')", "134-234");
ruby_test!(test_array_product_empty, "puts [1, 2].product([]).length", "0");
ruby_test!(test_array_product_block, "acc = []; [1, 2].product([3]) { |a, b| acc << a + b }; puts acc.join('-')", "4-5");
