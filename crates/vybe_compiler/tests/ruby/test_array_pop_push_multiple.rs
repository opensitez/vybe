
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_pop_multiple, "a = [1, 2, 3]; puts a.pop(2).join('-')", "2-3");
ruby_test!(test_pop_multiple_mutates, "a = [1, 2, 3]; a.pop(2); puts a.join('-')", "1");
ruby_test!(test_pop_more_than_length, "a = [1]; puts a.pop(3).join('-')", "1");
ruby_test!(test_pop_zero, "a = [1]; puts a.pop(0).length", "0");
ruby_test!(test_pop_empty_multiple, "a = []; puts a.pop(2).length", "0");
ruby_test!(test_push_multiple, "a = [1]; a.push(2, 3); puts a.join('-')", "1-2-3");
ruby_test!(test_push_returns_self, "a = [1]; puts a.push(2).object_id == a.object_id", "true");
ruby_test!(test_push_zero_args, "a = [1]; a.push(); puts a.join('-')", "1");
ruby_test!(test_push_to_empty, "a = []; a.push(1, 2); puts a.join('-')", "1-2");
ruby_test!(test_append_alias, "a = [1]; a.append(2); puts a.join('-')", "1-2");
ruby_test!(test_pop_negative_error, "a = [1]; begin; a.pop(-1); rescue ArgumentError; puts 'err'; end", "err");
