use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_frozen_literal_pragma, "# frozen_string_literal: true\ns = 'hello'; puts s.frozen?", "true");
ruby_test!(test_frozen_literal_mutation_error, "# frozen_string_literal: true\nbegin; 'hello' << 'world'; rescue => e; puts 'err'; end", "err");
ruby_test!(test_frozen_literal_object_id, "# frozen_string_literal: true\nputs 'a'.object_id == 'a'.object_id", "true");
ruby_test!(test_frozen_literal_interpolation, "# frozen_string_literal: true\nx = 1; s = \"a#{x}\"; puts s.frozen?", "false"); // Interpolated strings are not frozen by default
ruby_test!(test_frozen_literal_dynamic, "s = 'hello'.freeze; puts s.frozen?", "true");
ruby_test!(test_frozen_literal_dup, "# frozen_string_literal: true\ns = 'hello'.dup; puts s.frozen?", "false");
ruby_test!(test_frozen_literal_clone, "# frozen_string_literal: true\ns = 'hello'.clone; puts s.frozen?", "true"); // clone preserves frozen status
ruby_test!(test_frozen_literal_minus_string, "s = -'hello'; puts s.frozen?", "true"); // Dedup and freeze
ruby_test!(test_frozen_literal_plus_string, "s = +'hello'; puts s.frozen?", "false"); // Mutable copy
ruby_test!(test_frozen_literal_concat_result, "# frozen_string_literal: true\ns = 'a' + 'b'; puts s.frozen?", "false"); // Result of + is mutable
ruby_test!(test_frozen_literal_freeze_method, "s = 'a'; s.freeze; puts s.frozen?", "true");
ruby_test!(test_frozen_literal_array_of_strings, "# frozen_string_literal: true\na = ['a', 'b']; puts a[0].frozen?", "true");
