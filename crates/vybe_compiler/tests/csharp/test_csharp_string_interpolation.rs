use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

// ── String interpolation ($"...") ──────────────────────────────────────────

csharp_case!(
    interpolation_embeds_string_variable,
    r#"var name = "Alice"; Console.WriteLine($"{name} is here");"#,
    ["Alice is here"]
);
csharp_case!(
    interpolation_embeds_integer_variable,
    r#"var age = 30; Console.WriteLine($"Age: {age}");"#,
    ["Age: 30"]
);
csharp_case!(
    interpolation_evaluates_arithmetic_in_hole,
    r#"var a = 3; var b = 4; Console.WriteLine($"sum = {a + b}");"#,
    ["sum = 7"]
);
csharp_case!(
    interpolation_calls_method_inside_hole,
    r#"var name = "world"; Console.WriteLine($"Hello {name.ToUpper()}!");"#,
    ["Hello WORLD!"]
);
csharp_case!(
    interpolation_uses_ternary_inside_hole,
    r#"var x = 5; Console.WriteLine($"x is {(x > 3 ? "big" : "small")}");"#,
    ["x is big"]
);
csharp_case!(
    interpolation_escapes_literal_braces,
    r#"var val = 9; Console.WriteLine($"{{value}} = {val}");"#,
    ["{value} = 9"]
);
csharp_case!(
    interpolation_embeds_boolean_literal,
    r#"Console.WriteLine($"active={true}");"#,
    ["active=True"]
);
csharp_case!(
    interpolation_embeds_decimal_number,
    r#"var price = 3.5; Console.WriteLine($"price={price}");"#,
    ["price=3.5"]
);
csharp_case!(
    interpolation_with_empty_string_variable,
    r#"var s = ""; Console.WriteLine($"empty=[{s}]");"#,
    ["empty=[]"]
);
csharp_case!(
    interpolation_multiple_holes_in_sequence,
    r#"Console.WriteLine($"one{1}two{2}three");"#,
    ["one1two2three"]
);
csharp_case!(
    interpolation_expression_concatenates_inside_hole,
    r#"var s = "go"; Console.WriteLine($"next={s + "!"}");"#,
    ["next=go!"]
);
csharp_case!(
    interpolation_preserves_surrounding_whitespace,
    r#"var name = "Bob"; Console.WriteLine($" | {name} | ");"#,
    [" | Bob | "]
);
csharp_case!(
    interpolation_combined_with_plus_concatenation,
    r#"var name = "Ann"; Console.WriteLine($"Hi " + $"{name}");"#,
    ["Hi Ann"]
);

// ── Concatenation with + ───────────────────────────────────────────────────

csharp_case!(
    concat_two_string_literals,
    r#"Console.WriteLine("Hello" + " World");"#,
    ["Hello World"]
);
csharp_case!(
    concat_literal_then_variable,
    r#"var name = "Bob"; Console.WriteLine("Hello " + name);"#,
    ["Hello Bob"]
);
csharp_case!(
    concat_variable_then_literal,
    r#"var word = "go"; Console.WriteLine(word + "!");"#,
    ["go!"]
);
csharp_case!(
    concat_three_parts_with_spaces,
    r#"var first = "John"; var last = "Doe"; Console.WriteLine(first + " " + last);"#,
    ["John Doe"]
);
csharp_case!(
    concat_with_empty_string_leaves_value_unchanged,
    r#"Console.WriteLine("hello" + "");"#,
    ["hello"]
);
csharp_case!(
    concat_two_empty_strings_yields_empty,
    r#"Console.WriteLine("" + "");"#,
    [""]
);
csharp_case!(
    concat_integer_converted_via_to_string,
    r#"Console.WriteLine("n=" + 42.ToString());"#,
    ["n=42"]
);
csharp_case!(
    concat_after_string_method_call,
    r#"Console.WriteLine("  hi  ".Trim() + "!");"#,
    ["hi!"]
);

// ── Verbatim strings (@"") ─────────────────────────────────────────────────

csharp_case!(
    verbatim_preserves_backslashes_in_path,
    r#"var path = @"C:\Users\test\file.txt"; Console.WriteLine(path);"#,
    [r"C:\Users\test\file.txt"]
);
csharp_case!(
    verbatim_doubles_quotes_for_embedded_quote,
    r#"Console.WriteLine(@"say ""hello""");"#,
    [r#"say "hello""#]
);
csharp_case!(
    verbatim_does_not_interpret_backslash_n,
    r#"Console.WriteLine(@"line1\nline2");"#,
    [r"line1\nline2"]
);
csharp_case!(
    verbatim_empty_string_has_zero_length,
    r#"Console.WriteLine(@"".Length);"#,
    ["0"]
);
csharp_case!(
    verbatim_at_sign_in_content,
    r#"Console.WriteLine(@"user@host.com");"#,
    ["user@host.com"]
);
csharp_case!(
    verbatim_concatenated_with_regular_string,
    r#"Console.WriteLine(@"ab" + "cd");"#,
    ["abcd"]
);
csharp_case!(
    interpolated_verbatim_combines_formats,
    r#"var name = "docs"; Console.WriteLine($@"C:\{name}\readme.txt");"#,
    [r"C:\docs\readme.txt"]
);
csharp_case!(
    verbatim_stores_unicode_characters,
    r#"Console.WriteLine(@"café");"#,
    ["café"]
);

// ── Escape sequences ───────────────────────────────────────────────────────

csharp_case!(
    escape_backslash_renders_single_backslash,
    r#"Console.WriteLine("\\");"#,
    [r"\"]
);
csharp_case!(
    escape_double_quote_inside_double_quoted_string,
    r#"Console.WriteLine("\"");"#,
    [r#"""#]
);
csharp_case!(
    escape_newline_splits_output_line,
    r#"Console.WriteLine("a\nb");"#,
    ["a", "b"]
);
csharp_case!(
    escape_tab_inserts_horizontal_tab,
    r#"Console.WriteLine("a\tb");"#,
    ["a\tb"]
);
csharp_case!(
    escape_single_quote_allowed_in_double_quoted_string,
    r#"Console.WriteLine("'");"#,
    ["'"]
);
csharp_case!(
    escape_hex_sequence_renders_letter_a,
    r#"Console.WriteLine("\x41");"#,
    ["A"]
);
csharp_case!(
    escape_unicode_sequence_renders_letter_a,
    r#"Console.WriteLine("\u0041");"#,
    ["A"]
);

// ── Length ─────────────────────────────────────────────────────────────────

csharp_case!(
    length_counts_characters_in_word,
    r#"Console.WriteLine("hello".Length);"#,
    ["5"]
);
csharp_case!(
    length_of_empty_string_is_zero,
    r#"Console.WriteLine("".Length);"#,
    ["0"]
);
csharp_case!(
    length_counts_whitespace_characters,
    r#"Console.WriteLine("   ".Length);"#,
    ["3"]
);
csharp_case!(
    length_after_concatenation_reflects_total,
    r#"Console.WriteLine(("ab" + "cd").Length);"#,
    ["4"]
);
// ── ToUpper / ToLower ──────────────────────────────────────────────────────

csharp_case!(
    toupper_converts_lowercase_letters,
    r#"Console.WriteLine("hello".ToUpper());"#,
    ["HELLO"]
);
csharp_case!(
    toupper_leaves_already_uppercase_unchanged,
    r#"Console.WriteLine("HELLO".ToUpper());"#,
    ["HELLO"]
);
csharp_case!(
    tolower_converts_uppercase_letters,
    r#"Console.WriteLine("HELLO".ToLower());"#,
    ["hello"]
);
csharp_case!(
    tolower_normalizes_mixed_case,
    r#"Console.WriteLine("HeLLo".ToLower());"#,
    ["hello"]
);
csharp_case!(
    toupper_preserves_digit_characters,
    r#"Console.WriteLine("abc123".ToUpper());"#,
    ["ABC123"]
);

// ── Trim / TrimStart / TrimEnd ───────────────────────────────────────────────

csharp_case!(
    trim_removes_leading_and_trailing_spaces,
    r#"Console.WriteLine("  hello  ".Trim());"#,
    ["hello"]
);
csharp_case!(
    trim_leaves_string_without_whitespace_unchanged,
    r#"Console.WriteLine("hello".Trim());"#,
    ["hello"]
);
csharp_case!(
    trimstart_removes_leading_spaces_only,
    r#"Console.WriteLine("  hi".TrimStart());"#,
    ["hi"]
);
csharp_case!(
    trimend_removes_trailing_spaces_only,
    r#"Console.WriteLine("hi  ".TrimEnd());"#,
    ["hi"]
);
csharp_case!(
    trim_removes_tabs_and_spaces,
    r#"Console.WriteLine("\t hi \t".Trim());"#,
    ["hi"]
);
csharp_case!(
    trim_on_empty_string_returns_empty,
    r#"Console.WriteLine("".Trim());"#,
    [""]
);

// ── Substring ──────────────────────────────────────────────────────────────

csharp_case!(
    substring_from_index_to_end_of_string,
    r#"Console.WriteLine("hello world".Substring(6));"#,
    ["world"]
);
csharp_case!(
    substring_with_start_and_length,
    r#"Console.WriteLine("hello world".Substring(0, 5));"#,
    ["hello"]
);
csharp_case!(
    substring_single_character_slice,
    r#"Console.WriteLine("hello".Substring(1, 1));"#,
    ["e"]
);
csharp_case!(
    substring_from_start_index_zero,
    r#"Console.WriteLine("abcdef".Substring(0, 3));"#,
    ["abc"]
);
csharp_case!(
    substring_spanning_entire_string,
    r#"var s = "test"; Console.WriteLine(s.Substring(0, s.Length));"#,
    ["test"]
);
csharp_case!(
    substring_from_last_character,
    r#"Console.WriteLine("abc".Substring(2));"#,
    ["c"]
);

// ── IndexOf ────────────────────────────────────────────────────────────────

csharp_case!(
    indexof_finds_substring_at_beginning,
    r#"Console.WriteLine("hello".IndexOf("he"));"#,
    ["0"]
);
csharp_case!(
    indexof_finds_substring_in_middle,
    r#"Console.WriteLine("hello world".IndexOf("world"));"#,
    ["6"]
);
csharp_case!(
    indexof_returns_negative_one_when_missing,
    r#"Console.WriteLine("hello".IndexOf("xyz"));"#,
    ["-1"]
);
csharp_case!(
    indexof_finds_single_character,
    r#"Console.WriteLine("banana".IndexOf("a"));"#,
    ["1"]
);
csharp_case!(
    indexof_with_start_index_skips_earlier_match,
    r#"Console.WriteLine("hello".IndexOf("l", 2));"#,
    ["3"]
);

// ── Contains ───────────────────────────────────────────────────────────────

csharp_case!(
    contains_reports_true_for_present_substring,
    r#"Console.WriteLine("hello world".Contains("world"));"#,
    ["True"]
);
csharp_case!(
    contains_reports_false_for_absent_substring,
    r#"Console.WriteLine("hello world".Contains("xyz"));"#,
    ["False"]
);
csharp_case!(
    contains_empty_substring_is_always_true,
    r#"Console.WriteLine("hello".Contains(""));"#,
    ["True"]
);
csharp_case!(
    contains_is_case_sensitive,
    r#"Console.WriteLine("Hello".Contains("hello"));"#,
    ["False"]
);

// ── StartsWith / EndsWith ──────────────────────────────────────────────────

csharp_case!(
    startswith_reports_true_for_matching_prefix,
    r#"Console.WriteLine("hello".StartsWith("hel"));"#,
    ["True"]
);
csharp_case!(
    startswith_reports_false_for_non_matching_prefix,
    r#"Console.WriteLine("hello".StartsWith("xyz"));"#,
    ["False"]
);
csharp_case!(
    endswith_reports_true_for_matching_suffix,
    r#"Console.WriteLine("hello".EndsWith("llo"));"#,
    ["True"]
);
csharp_case!(
    endswith_reports_false_for_non_matching_suffix,
    r#"Console.WriteLine("hello".EndsWith("xyz"));"#,
    ["False"]
);
// ── Replace ────────────────────────────────────────────────────────────────

csharp_case!(
    replace_substitutes_all_matching_occurrences,
    r#"Console.WriteLine("hello".Replace("l", "L"));"#,
    ["heLLo"]
);
csharp_case!(
    replace_single_occurrence_in_word,
    r#"Console.WriteLine("hello world".Replace("world", "C#"));"#,
    ["hello C#"]
);
csharp_case!(
    replace_with_no_match_returns_original,
    r#"Console.WriteLine("hello".Replace("xyz", "abc"));"#,
    ["hello"]
);
csharp_case!(
    replace_with_empty_string_removes_matches,
    r#"Console.WriteLine("aba".Replace("a", ""));"#,
    ["b"]
);
csharp_case!(
    replace_entire_string_with_new_value,
    r#"Console.WriteLine("old".Replace("old", "new"));"#,
    ["new"]
);

// ── Split ──────────────────────────────────────────────────────────────────

csharp_case!(
    split_comma_delimited_string_into_parts,
    r#"var parts = "a,b,c".Split(","); Console.WriteLine(parts.Length); Console.WriteLine(parts[1]);"#,
    ["3", "b"]
);
csharp_case!(
    split_without_delimiter_yields_single_element,
    r#"var parts = "solo".Split(","); Console.WriteLine(parts.Length);"#,
    ["1"]
);
csharp_case!(
    split_preserves_empty_field_between_delimiters,
    r#"var parts = "a,,b".Split(","); Console.WriteLine(parts[1]);"#,
    [""]
);
csharp_case!(
    split_on_space_separates_words,
    r#"var parts = "one two".Split(" "); Console.WriteLine(parts[0]); Console.WriteLine(parts[1]);"#,
    ["one", "two"]
);
csharp_case!(
    split_first_element_of_delimited_list,
    r#"var parts = "x-y-z".Split("-"); Console.WriteLine(parts[0]);"#,
    ["x"]
);

// ── Join ───────────────────────────────────────────────────────────────────

csharp_case!(
    join_combines_array_elements_with_delimiter,
    r#"Console.WriteLine(string.Join("|", new[] { "a", "b", "c" }));"#,
    ["a|b|c"]
);
csharp_case!(
    join_empty_array_produces_empty_string,
    r#"Console.WriteLine(string.Join(",", new string[] { }));"#,
    [""]
);
csharp_case!(
    join_single_element_without_extra_delimiter,
    r#"Console.WriteLine(string.Join("-", new[] { "only" }));"#,
    ["only"]
);
csharp_case!(
    join_with_space_separator,
    r#"Console.WriteLine(string.Join(" ", new[] { "Hello", "World" }));"#,
    ["Hello World"]
);

// ── IsNullOrEmpty ──────────────────────────────────────────────────────────

csharp_case!(
    isnullorempty_reports_true_for_null,
    r#"Console.WriteLine(string.IsNullOrEmpty(null));"#,
    ["True"]
);
csharp_case!(
    isnullorempty_reports_true_for_empty_string,
    r#"Console.WriteLine(string.IsNullOrEmpty(""));"#,
    ["True"]
);
csharp_case!(
    isnullorempty_reports_false_for_nonempty_string,
    r#"Console.WriteLine(string.IsNullOrEmpty("hello"));"#,
    ["False"]
);
csharp_case!(
    isnullorempty_reports_false_for_whitespace_only,
    r#"Console.WriteLine(string.IsNullOrEmpty(" "));"#,
    ["False"]
);

// ── PadLeft / PadRight ─────────────────────────────────────────────────────

csharp_case!(
    padleft_zero_fills_leading_characters,
    r#"Console.WriteLine("5".PadLeft(3, '0'));"#,
    ["005"]
);
csharp_case!(
    padright_zero_fills_trailing_characters,
    r#"Console.WriteLine("5".PadRight(3, '0'));"#,
    ["500"]
);
csharp_case!(
    padleft_when_already_long_enough_returns_original,
    r#"Console.WriteLine("hello".PadLeft(3, '0'));"#,
    ["hello"]
);
csharp_case!(
    padright_with_space_character,
    r#"Console.WriteLine("x".PadRight(4));"#,
    ["x   "]
);
csharp_case!(
    padleft_to_exact_length_adds_no_padding,
    r#"Console.WriteLine("ab".PadLeft(2, '0'));"#,
    ["ab"]
);

// ── Insert ─────────────────────────────────────────────────────────────────

csharp_case!(
    insert_adds_text_at_specified_index,
    r#"Console.WriteLine("hello".Insert(2, "XX"));"#,
    ["heXXllo"]
);
csharp_case!(
    insert_at_beginning_prefixes_text,
    r#"Console.WriteLine("world".Insert(0, "Hello "));"#,
    ["Hello world"]
);
csharp_case!(
    insert_at_end_appends_text,
    r#"Console.WriteLine("end".Insert(3, "!"));"#,
    ["end!"]
);

// ── Remove ─────────────────────────────────────────────────────────────────

csharp_case!(
    remove_deletes_characters_starting_at_index,
    r#"Console.WriteLine("hello".Remove(1, 3));"#,
    ["ho"]
);
csharp_case!(
    remove_from_start_truncates_prefix,
    r#"Console.WriteLine("abcdef".Remove(0, 2));"#,
    ["cdef"]
);
csharp_case!(
    remove_from_index_to_end,
    r#"Console.WriteLine("hello".Remove(3));"#,
    ["hel"]
);
