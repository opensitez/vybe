
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerator_yielder_basic, "enum = Enumerator.new { |y| y << 1; y << 2 }; puts enum.to_a.join('-')", "1-2");
ruby_test!(test_enumerator_yielder_yield, "enum = Enumerator.new { |y| y.yield(1); y.yield(2) }; puts enum.to_a.join('-')", "1-2");
ruby_test!(test_enumerator_yielder_multiple, "enum = Enumerator.new { |y| y.yield(1, 2) }; puts enum.to_a.flatten.join('-')", "1-2");
ruby_test!(test_enumerator_generator_basic, "enum = Enumerator.new { |y| y << 1 }; puts enum.next", "1");
ruby_test!(test_enumerator_generator_stop, "enum = Enumerator.new { |y| y << 1 }; enum.next; begin; enum.next; rescue StopIteration; puts 'err'; end", "err");
