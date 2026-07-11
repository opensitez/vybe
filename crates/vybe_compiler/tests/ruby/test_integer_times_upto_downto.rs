
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_integer_times_basic, "acc = []; 3.times { |i| acc << i }; puts acc.join('-')", "0-1-2");
ruby_test!(test_integer_times_zero, "acc = []; 0.times { |i| acc << i }; puts acc.join('-')", "");
ruby_test!(test_integer_times_negative, "acc = []; -3.times { |i| acc << i }; puts acc.join('-')", "");
ruby_test!(test_integer_upto_basic, "acc = []; 1.upto(3) { |i| acc << i }; puts acc.join('-')", "1-2-3");
ruby_test!(test_integer_upto_same, "acc = []; 3.upto(3) { |i| acc << i }; puts acc.join('-')", "3");
ruby_test!(test_integer_upto_less, "acc = []; 3.upto(1) { |i| acc << i }; puts acc.join('-')", "");
ruby_test!(test_integer_downto_basic, "acc = []; 3.downto(1) { |i| acc << i }; puts acc.join('-')", "3-2-1");
ruby_test!(test_integer_downto_same, "acc = []; 3.downto(3) { |i| acc << i }; puts acc.join('-')", "3");
ruby_test!(test_integer_downto_greater, "acc = []; 1.downto(3) { |i| acc << i }; puts acc.join('-')", "");
ruby_test!(test_integer_times_no_block, "puts 3.times.class.name", "Enumerator");
ruby_test!(test_integer_upto_no_block, "puts 1.upto(3).class.name", "Enumerator");
ruby_test!(test_integer_downto_no_block, "puts 3.downto(1).class.name", "Enumerator");
