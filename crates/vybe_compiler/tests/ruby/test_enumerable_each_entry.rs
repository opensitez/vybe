use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_each_entry_basic, "acc = []; [1, 2].each_entry {|x| acc << x}; puts acc.join('-')", "1-2");
ruby_test!(test_each_entry_no_block, "puts [1, 2].each_entry.is_a?(Enumerator)", "true");
ruby_test!(test_each_entry_hash, "acc = []; {a: 1}.each_entry {|kv| acc << kv.join('-')}; puts acc.join('-')", "a-1");
ruby_test!(test_each_entry_yields_array, "class A; include Enumerable; def each; yield 1, 2; end; end; acc = []; A.new.each_entry {|x| acc << x.inspect}; puts acc.join('-')", "[1, 2]"); // each_entry wraps multiple yielded values in an array
ruby_test!(test_each_entry_yields_single, "class A; include Enumerable; def each; yield 1; end; end; acc = []; A.new.each_entry {|x| acc << x.inspect}; puts acc.join('-')", "1");
ruby_test!(test_each_entry_returns_self, "a = [1]; puts a.each_entry {|x| x}.object_id == a.object_id", "true");
ruby_test!(test_each_entry_empty, "acc = []; [].each_entry {|x| acc << x}; puts acc.length", "0");
