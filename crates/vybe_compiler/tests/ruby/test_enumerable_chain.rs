
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_chain_basic, "puts [1, 2].chain([3, 4]).to_a.join('-')", "1-2-3-4");
ruby_test!(test_chain_multiple, "puts [1].chain([2], [3]).to_a.join('-')", "1-2-3");
ruby_test!(test_chain_no_args, "puts [1, 2].chain.to_a.join('-')", "1-2");
ruby_test!(test_chain_returns_enumerator_chain, "puts [1, 2].chain([3, 4]).class.name", "Enumerator::Chain");
ruby_test!(test_chain_empty, "puts [].chain([1]).to_a.join('-')", "1");
ruby_test!(test_chain_all_empty, "puts [].chain([]).to_a.length", "0");
ruby_test!(test_chain_non_array, "puts [1].chain(2..3).to_a.join('-')", "1-2-3");
ruby_test!(test_chain_iteration, "acc = []; [1].chain([2]).each {|x| acc << x}; puts acc.join('-')", "1-2");
ruby_test!(test_chain_enum_chain_method, "puts Enumerator::Chain.new([1], [2]).to_a.join('-')", "1-2");
