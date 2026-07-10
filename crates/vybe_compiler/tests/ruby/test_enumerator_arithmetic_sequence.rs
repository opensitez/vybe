use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerator_arithmetic_sequence_basic, "puts 1.step(10, 2).class.name", "Enumerator::ArithmeticSequence");
ruby_test!(test_enumerator_arithmetic_sequence_begin, "puts 1.step(10, 2).begin", "1");
ruby_test!(test_enumerator_arithmetic_sequence_end, "puts 1.step(10, 2).end", "10");
ruby_test!(test_enumerator_arithmetic_sequence_step, "puts 1.step(10, 2).step", "2");
ruby_test!(test_enumerator_arithmetic_sequence_exclude_end, "puts 1.step(10, 2).exclude_end?", "false"); // wait, step doesn't exclude end? It returns true if it came from Range with exclude_end.
ruby_test!(test_enumerator_arithmetic_sequence_from_range, "puts (1...10).step(2).exclude_end?", "true");
ruby_test!(test_enumerator_arithmetic_sequence_size, "puts 1.step(10, 2).size", "5");
ruby_test!(test_enumerator_arithmetic_sequence_first, "puts 1.step(10, 2).first", "1");
ruby_test!(test_enumerator_arithmetic_sequence_first_n, "puts 1.step(10, 2).first(2).join('-')", "1-3");
ruby_test!(test_enumerator_arithmetic_sequence_last, "puts 1.step(10, 2).last", "9");
ruby_test!(test_enumerator_arithmetic_sequence_last_n, "puts 1.step(10, 2).last(2).join('-')", "7-9");
ruby_test!(test_enumerator_arithmetic_sequence_hash, "puts 1.step(10, 2).hash.class.name", "Integer");
ruby_test!(test_enumerator_arithmetic_sequence_eq, "puts 1.step(10, 2) == 1.step(10, 2)", "true");
