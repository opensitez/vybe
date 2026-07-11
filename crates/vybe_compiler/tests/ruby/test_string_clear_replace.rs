
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_clear_basic, "s = 'hello'; s.clear; puts s", "");
ruby_test!(test_clear_returns_self, "s = 'hello'; puts s.clear.object_id == s.object_id", "true");
ruby_test!(test_replace_basic, "s = 'hello'; s.replace('world'); puts s", "world");
ruby_test!(test_replace_returns_self, "s = 'hello'; puts s.replace('world').object_id == s.object_id", "true");
ruby_test!(test_replace_changes_length, "s = 'a'; s.replace('abc'); puts s.length", "3");
ruby_test!(test_replace_preserves_object_id, "s = 'a'; id = s.object_id; s.replace('b'); puts s.object_id == id", "true");
ruby_test!(test_replace_with_empty, "s = 'a'; s.replace(''); puts s", "");
ruby_test!(test_clear_already_empty, "s = ''; s.clear; puts s", "");
ruby_test!(test_replace_same_string, "s = 'a'; s.replace(s); puts s", "a");
ruby_test!(test_replace_frozen_error, "# frozen_string_literal: true\ns = 'a'; begin; s.replace('b'); rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_clear_frozen_error, "# frozen_string_literal: true\ns = 'a'; begin; s.clear; rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_replace_encoding, "s = 'a'; s.replace('b'.force_encoding('UTF-8')); puts s.encoding.name", "UTF-8");
