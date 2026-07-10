use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_group_by_basic, "h = [1, 2, 3, 4].group_by {|x| x % 2}; puts h[0].join('-')", "2-4");
ruby_test!(test_group_by_keys, "h = [1, 2, 3, 4].group_by {|x| x % 2}; puts h.keys.sort.join('-')", "0-1");
ruby_test!(test_group_by_no_block, "puts [1].group_by.is_a?(Enumerator)", "true");
ruby_test!(test_group_by_empty, "puts [].group_by {|x| x}.length", "0");
ruby_test!(test_group_by_hash, "h = {a: 1, b: 2, c: 1}.group_by {|k, v| v}; puts h[1].map{|k, v| k.to_s}.sort.join('-')", "a-c");
ruby_test!(test_tally_basic, "h = ['a', 'b', 'a'].tally; puts h['a']", "2");
ruby_test!(test_tally_missing, "h = ['a', 'b'].tally; puts h['c'].nil?", "true"); // ruby 2.7+
ruby_test!(test_tally_empty, "puts [].tally.length", "0");
ruby_test!(test_tally_hash_basic, "puts {a: 1, b: 1}.tally.values.join('-')", "1-1"); // elements of hash are pairs, so each pair is unique usually unless duplicates
ruby_test!(test_tally_with_hash_arg, "h = {'a' => 1}; ['a', 'b'].tally(h); puts h['a']", "2"); // ruby 3.1+ tally(hash)
