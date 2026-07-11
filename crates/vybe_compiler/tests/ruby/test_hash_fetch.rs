
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_fetch_basic, "puts ({a: 1}.fetch(:a))", "1");
ruby_test!(test_hash_fetch_missing, "begin; {a: 1}.fetch(:b); rescue KeyError; puts 'err'; end", "err");
ruby_test!(test_hash_fetch_default, "puts ({a: 1}.fetch(:b, 2))", "2");
ruby_test!(test_hash_fetch_block, "puts ({a: 1}.fetch(:b) { |k| \"missing #{k}\" })", "missing b");
ruby_test!(test_hash_fetch_block_and_default, "puts ({a: 1}.fetch(:b, 2) { |k| 3 })", "3"); // block takes precedence over default arg (ruby gives warning but runs block)
ruby_test!(test_hash_fetch_values, "puts ({a: 1, b: 2, c: 3}.fetch_values(:a, :c).join('-'))", "1-3");
ruby_test!(test_hash_fetch_values_missing, "begin; {a: 1}.fetch_values(:a, :b); rescue KeyError; puts 'err'; end", "err");
ruby_test!(test_hash_fetch_values_block, "puts ({a: 1}.fetch_values(:a, :b) { |k| k == :b ? 2 : 0 }.join('-'))", "1-2");
