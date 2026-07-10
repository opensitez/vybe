use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_try_convert_hash, "puts Hash.try_convert({a: 1}).is_a?(Hash)", "true");
ruby_test!(test_try_convert_object, "class A; def to_hash; {a: 1}; end; end; puts Hash.try_convert(A.new).is_a?(Hash)", "true");
ruby_test!(test_try_convert_object_nil, "class A; def to_hash; nil; end; end; puts Hash.try_convert(A.new).nil?", "true");
ruby_test!(test_try_convert_object_error, "class A; def to_hash; 5; end; end; begin; Hash.try_convert(A.new); rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_try_convert_nil, "puts Hash.try_convert(nil).nil?", "true");
ruby_test!(test_try_convert_array, "puts Hash.try_convert([]).nil?", "true"); // does not convert arrays
ruby_test!(test_try_convert_string, "puts Hash.try_convert('a').nil?", "true");
ruby_test!(test_to_h_basic, "puts {a: 1}.to_h.is_a?(Hash)", "true");
ruby_test!(test_to_h_returns_self, "h = {a: 1}; puts h.to_h.object_id == h.object_id", "true");
ruby_test!(test_to_h_with_block, "puts {a: 1, b: 2}.to_h {|k, v| [k.to_s, v * 2]}['b']", "4"); // ruby 2.6+
ruby_test!(test_to_h_with_block_error, "begin; {a: 1}.to_h {|k, v| 5}; rescue TypeError; puts 'err'; end", "err"); // block must return array of 2 elements
