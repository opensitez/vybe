use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_each_with_index_basic, "acc = []; [1, 2].each_with_index {|x, i| acc << \"#{x}:#{i}\"}; puts acc.join('-')", "1:0-2:1");
ruby_test!(test_each_with_index_hash, "acc = []; {a: 1}.each_with_index {|kv, i| acc << \"#{kv[0]}:#{i}\"}; puts acc.join('-')", "a:0");
ruby_test!(test_each_with_index_no_block, "puts [1].each_with_index.is_a?(Enumerator)", "true");
ruby_test!(test_each_with_index_returns_self, "a = [1]; puts a.each_with_index {|x, i|}.object_id == a.object_id", "true");
ruby_test!(test_each_with_object_basic, "puts [1, 2].each_with_object([]) {|x, o| o << x * 2}.join('-')", "2-4");
ruby_test!(test_each_with_object_hash, "puts {a: 1, b: 2}.each_with_object({}) {|kv, o| o[kv[0]] = kv[1] * 2}[:b]", "4");
ruby_test!(test_each_with_object_no_block, "puts [1].each_with_object([]).is_a?(Enumerator)", "true");
ruby_test!(test_each_with_object_returns_object, "o = []; puts [1].each_with_object(o) {|x, ob|}.object_id == o.object_id", "true");
ruby_test!(test_each_with_index_empty, "acc = []; [].each_with_index {|x, i| acc << i}; puts acc.length", "0");
ruby_test!(test_each_with_object_empty, "puts [].each_with_object([1]) {|x, o| o << 2}.join('-')", "1");
