
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_any_basic, "puts ({a: 1, b: 2}.any?)", "true");
ruby_test!(test_any_empty, "puts ({}).any?", "false");
ruby_test!(test_any_block, "puts ({a: 1, b: 2}.any? {|k, v| v > 1})", "true");
ruby_test!(test_any_block_false, "puts ({a: 1, b: 2}.any? {|k, v| v > 5})", "false");
ruby_test!(test_any_argument, "puts ({a: 1, b: 2}.any?([:a, 1]))", "true"); // argument matches pair
ruby_test!(test_any_argument_false, "puts ({a: 1, b: 2}.any?([:a, 2]))", "false");
ruby_test!(test_all_basic, "puts ({a: 1, b: 2}.all?)", "true"); // pairs evaluate to true
ruby_test!(test_all_empty, "puts ({}).all?", "true"); // vacuously true
ruby_test!(test_all_block, "puts ({a: 1, b: 2}.all? {|k, v| v > 0})", "true");
ruby_test!(test_all_block_false, "puts ({a: 1, b: 2}.all? {|k, v| v > 1})", "false");
ruby_test!(test_none_basic, "puts ({a: 1}.none?)", "false");
ruby_test!(test_none_empty, "puts ({}).none?", "true");
ruby_test!(test_none_block, "puts ({a: 1, b: 2}.none? {|k, v| v > 5})", "true");
ruby_test!(test_none_block_false, "puts ({a: 1, b: 2}.none? {|k, v| v > 1})", "false");
ruby_test!(test_one_basic, "puts ({a: 1}.one?)", "true");
ruby_test!(test_one_multiple, "puts ({a: 1, b: 2}.one?)", "false");
ruby_test!(test_one_empty, "puts ({}).one?", "false");
ruby_test!(test_one_block, "puts ({a: 1, b: 2}.one? {|k, v| v > 1})", "true");
ruby_test!(test_one_block_false, "puts ({a: 1, b: 2}.one? {|k, v| v > 0})", "false");
