use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_manipulation_push, "a = [1]; a.push(2, 3); puts a.join('-')", "1-2-3");
ruby_test!(test_array_manipulation_pop, "a = [1, 2]; puts a.pop", "2");
ruby_test!(test_array_manipulation_pop_multiple, "a = [1, 2, 3]; puts a.pop(2).join('-')", "2-3");
ruby_test!(test_array_manipulation_shift, "a = [1, 2]; puts a.shift", "1");
ruby_test!(test_array_manipulation_shift_multiple, "a = [1, 2, 3]; puts a.shift(2).join('-')", "1-2");
ruby_test!(test_array_manipulation_unshift, "a = [1]; a.unshift(2, 3); puts a.join('-')", "2-3-1");
ruby_test!(test_array_manipulation_insert, "a = [1, 2]; a.insert(1, 3, 4); puts a.join('-')", "1-3-4-2");
ruby_test!(test_array_manipulation_insert_negative, "a = [1, 2]; a.insert(-1, 3); puts a.join('-')", "1-2-3");
ruby_test!(test_array_manipulation_delete, "a = [1, 2, 1]; a.delete(1); puts a.join('-')", "2");
ruby_test!(test_array_manipulation_delete_block, "a = [1, 2]; puts a.delete(3) { 'not found' }", "not found");
ruby_test!(test_array_manipulation_delete_at, "a = [1, 2, 3]; puts a.delete_at(1)", "2");
ruby_test!(test_array_manipulation_delete_if, "a = [1, 2, 3]; a.delete_if { |x| x.even? }; puts a.join('-')", "1-3");
ruby_test!(test_array_manipulation_keep_if, "a = [1, 2, 3]; a.keep_if { |x| x.even? }; puts a.join('-')", "2");
