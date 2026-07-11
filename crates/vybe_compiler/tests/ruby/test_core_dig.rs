
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_dig, "h = { a: { b: { c: 1 } } }; puts h.dig(:a, :b, :c)", "1");
ruby_test!(test_hash_dig_missing, "h = { a: { b: 1 } }; puts h.dig(:a, :c).nil?", "true");
ruby_test!(test_hash_dig_not_hash, "h = { a: 1 }; begin; h.dig(:a, :b); rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_array_dig, "a = [[1, [2, 3]]]; puts a.dig(0, 1, 1)", "3");
ruby_test!(test_array_dig_missing, "a = [[1, 2]]; puts a.dig(0, 5).nil?", "true");
ruby_test!(test_struct_dig, "S = Struct.new(:a); s = S.new({b: 1}); puts s.dig(:a, :b)", "1");
