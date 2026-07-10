use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_heredoc_basic, "puts <<EOF\nhello\nEOF\n", "hello");
ruby_test!(test_heredoc_indented, "puts <<-EOF\n  hello\nEOF\n", "  hello");
ruby_test!(test_heredoc_squiggly, "puts <<~EOF\n  hello\nEOF\n", "hello");
ruby_test!(test_heredoc_multiple, "puts <<A, <<B\none\nA\ntwo\nB\n", "one\ntwo");
ruby_test!(test_heredoc_single_quotes, "name = 'x'; puts <<'EOF'\n#{name}\nEOF\n", "#{name}");
ruby_test!(test_heredoc_double_quotes, "name = 'world'; puts <<\"EOF\"\nhello #{name}\nEOF\n", "hello world");
ruby_test!(test_heredoc_backticks, "puts <<`EOF`\necho hello\nEOF\n", "hello"); // Assuming system command works, else just parsing
ruby_test!(test_heredoc_empty, "puts <<EOF\nEOF\n", "");
ruby_test!(test_heredoc_method_call, "puts <<EOF.upcase\nhello\nEOF\n", "HELLO");
ruby_test!(test_heredoc_stacked_method, "puts <<EOF.strip.upcase\n  hello  \nEOF\n", "HELLO");
ruby_test!(test_heredoc_in_array, "a = [<<A, <<B]\none\nA\ntwo\nB\nputs a[0]", "one\n");
ruby_test!(test_heredoc_interpolated_math, "puts <<EOF\n#{1 + 1}\nEOF\n", "2");
