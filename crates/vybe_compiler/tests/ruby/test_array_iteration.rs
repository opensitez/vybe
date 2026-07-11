
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_iteration_each, "acc = []; [1, 2].each { |x| acc << x }; puts acc.join('-')", "1-2");
ruby_test!(test_array_iteration_each_with_index, "acc = []; [1, 2].each_with_index { |x, i| acc << \"#{x}:#{i}\" }; puts acc.join('-')", "1:0-2:1");
ruby_test!(test_array_iteration_reverse_each, "acc = []; [1, 2].reverse_each { |x| acc << x }; puts acc.join('-')", "2-1");
ruby_test!(test_array_iteration_cycle, "acc = []; [1, 2].cycle(2) { |x| acc << x }; puts acc.join('-')", "1-2-1-2");
ruby_test!(test_array_iteration_each_enumerator, "puts [1].each.class.name", "Enumerator");
ruby_test!(test_array_iteration_each_with_index_enumerator, "puts [1].each_with_index.class.name", "Enumerator");
ruby_test!(test_array_iteration_reverse_each_enumerator, "puts [1].reverse_each.class.name", "Enumerator");
ruby_test!(test_array_iteration_cycle_enumerator, "puts [1].cycle.class.name", "Enumerator");
ruby_test!(test_array_iteration_each_return, "puts [1, 2].each { |x| x }.join('-')", "1-2"); // each returns self
