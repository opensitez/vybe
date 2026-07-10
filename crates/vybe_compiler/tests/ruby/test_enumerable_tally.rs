use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_tally_basic, "puts ['a', 'b', 'a'].tally['a']", "2");
ruby_test!(test_enumerable_tally_keys, "puts ['a', 'b', 'a'].tally.keys.sort.join('-')", "a-b");
ruby_test!(test_enumerable_tally_empty, "puts [].tally.length", "0");
ruby_test!(test_enumerable_tally_hash, "h = {a: 0}; ['a', 'b', 'a'].tally(h); puts h['a']", "2"); // ruby 2.7 added tally, 3.0 added tally(hash)
ruby_test!(test_enumerable_tally_hash_init, "h = {'a' => 10}; ['a', 'b', 'a'].tally(h); puts h['a']", "12");
ruby_test!(test_enumerable_tally_mixed, "puts [1, '1', 1].tally[1]", "2");
