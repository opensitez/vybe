use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dig_basic, "a = [[1, [2, 3]]]; puts a.dig(0, 1, 1)", "3");
ruby_test!(test_dig_first_level, "a = [1, 2]; puts a.dig(1)", "2");
ruby_test!(test_dig_missing_index, "a = [[1]]; puts a.dig(1, 0).nil?", "true");
ruby_test!(test_dig_missing_deep_index, "a = [[1]]; puts a.dig(0, 2).nil?", "true");
ruby_test!(test_dig_type_error, "a = [[1]]; begin; a.dig(0, 0, 0); rescue TypeError; puts 'err'; end", "err"); // 1 does not have dig method
ruby_test!(test_dig_hash_mix, "a = [{key: [1, 2]}]; puts a.dig(0, :key, 1)", "2");
ruby_test!(test_dig_no_args_error, "a = [1]; begin; a.dig(); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_dig_negative_index, "a = [[1, 2], [3, 4]]; puts a.dig(-1, -1)", "4");
ruby_test!(test_dig_missing_negative_index, "a = [[1]]; puts a.dig(-2, 0).nil?", "true");
ruby_test!(test_dig_struct_mix, "S = Struct.new(:a); a = [S.new([1, 2])]; puts a.dig(0, :a, 1)", "2");
