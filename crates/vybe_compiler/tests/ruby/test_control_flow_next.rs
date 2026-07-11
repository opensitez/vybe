
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_next_basic, "acc = []; for i in 1..3; next if i == 2; acc << i; end; puts acc.join('-')", "1-3");
ruby_test!(test_next_value, "acc = []; acc << (1..2).map { |i| next 'val' if i == 1; i }.join('-'); puts acc.join", "val-2"); // next returns value from block
ruby_test!(test_next_while, "i = 0; acc = []; while i < 3; i += 1; next if i == 2; acc << i; end; puts acc.join('-')", "1-3");
ruby_test!(test_next_nested, "acc = []; for i in 1..2; for j in 1..2; next if j == 1; acc << \"#{i}-#{j}\"; end; end; puts acc.join('|')", "1-2|2-2");
ruby_test!(test_next_error, "begin; eval('next'); rescue SyntaxError; puts 'err'; end", "err");
