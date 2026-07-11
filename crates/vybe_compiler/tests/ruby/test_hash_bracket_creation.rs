
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_square_brackets_empty, "puts Hash[].length", "0");
ruby_test!(test_square_brackets_pairs, "puts Hash['a', 1, 'b', 2]['a']", "1");
ruby_test!(test_square_brackets_odd_error, "begin; Hash['a', 1, 'b']; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_square_brackets_array_of_pairs, "puts Hash[[['a', 1], ['b', 2]]]['b']", "2");
ruby_test!(test_square_brackets_array_of_pairs_mixed, "puts Hash[[['a', 1], ['b', 2, 3]]]['b']", "2"); // Wait, Array with 3 elements? Hash[[['b', 2, 3]]] raises ArgumentError usually
ruby_test!(test_square_brackets_array_of_pairs_error, "begin; Hash[[['a', 1], ['b', 2, 3]]]; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_square_brackets_array_of_pairs_error_1, "begin; Hash[[['a', 1], ['b']]]; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_square_brackets_hash, "puts Hash[{a: 1}][:a]", "1");
ruby_test!(test_square_brackets_object_to_hash, "class A; def to_hash; {a: 1}; end; end; puts Hash[A.new][:a]", "1");
ruby_test!(test_new_basic, "puts Hash.new.length", "0");
ruby_test!(test_new_default_value, "h = Hash.new(5); puts h[:a]", "5");
ruby_test!(test_new_default_value_same_object, "h = Hash.new([]); h[:a] << 1; h[:b] << 2; puts h[:c].join('-')", "1-2");
ruby_test!(test_new_default_block, "h = Hash.new {|hash, key| hash[key] = key.to_s}; puts h[:a]", "a");
ruby_test!(test_new_default_block_mutates, "h = Hash.new {|hash, key| hash[key] = key.to_s}; h[:a]; puts h.length", "1");
ruby_test!(test_new_with_arguments_and_block_error, "begin; Hash.new(5) {|hash, key| 1}; rescue ArgumentError; puts 'err'; end", "err"); // can't provide both
