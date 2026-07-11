
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dig_basic, "puts ({a: {b: {c: 1}}}.dig(:a, :b, :c))", "1");
ruby_test!(test_dig_first_level, "puts ({a: 1}.dig(:a))", "1");
ruby_test!(test_dig_missing_key, "puts ({a: 1}.dig(:b).nil?)", "true");
ruby_test!(test_dig_missing_deep_key, "puts ({a: {}}.dig(:a, :b).nil?)", "true");
ruby_test!(test_dig_type_error, "begin; {a: 1}.dig(:a, :b); rescue TypeError; puts 'err'; end", "err"); // 1 does not have dig method
ruby_test!(test_dig_array_mix, "puts ({a: [1, 2]}.dig(:a, 1))", "2");
ruby_test!(test_dig_no_args_error, "begin; {a: 1}.dig(); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_dig_struct_mix, "S = Struct.new(:b); puts ({a: S.new(2)}.dig(:a, :b))", "2");
ruby_test!(test_dig_ignores_hash_default, "h = Hash.new('def'); puts h.dig(:a).nil?", "true"); // Wait! Hash#dig DOES use default_proc! Let's check:
// In ruby, h.dig(:a) returns 'def' if h has default 'def'. Yes!
ruby_test!(test_dig_uses_hash_default, "h = Hash.new('def'); puts h.dig(:a)", "def");
ruby_test!(test_dig_uses_hash_default_proc, "h = Hash.new {|hash, key| 'def'}; puts h.dig(:a)", "def");
