//! Raw string literals (`"""`), interpolation (`$"""`), and multiline forms.
//! GAP: raw-string coverage is thin beyond a single verbatim/raw smoke test.

csharp_cases! {
    raw_string_basic_content_without_escapes => {
        r####"string text="""hello"""; Console.WriteLine(text);"####,
        ["hello"]
    };

    raw_string_embedded_double_quotes_without_backslash => {
        r####"string text="""say "hi" now"""; Console.WriteLine(text.Contains("\"hi\""));"####,
        ["True"]
    };

    raw_string_multiline_preserves_internal_newline => {
        r####"string text="""line1
line2"""; Console.WriteLine(text.Contains("\n"));"####,
        ["True"]
    };

    raw_string_multiline_first_line_content => {
        r####"string text="""alpha
beta"""; Console.WriteLine(text.StartsWith("alpha"));"####,
        ["True"]
    };

    raw_string_multiline_second_line_content => {
        r####"string text="""alpha
beta"""; Console.WriteLine(text.EndsWith("beta"));"####,
        ["True"]
    };

    raw_string_empty_content_has_zero_length => {
        r####"string text="""""""; Console.WriteLine(text.Length);"####,
        ["0"]
    };

    raw_string_single_character_length => {
        r####"string text="""x"""; Console.WriteLine(text.Length);"####,
        ["1"]
    };

    raw_string_leading_spaces_are_preserved => {
        r####"string text="""  spaced"""; Console.WriteLine(text.StartsWith("  "));"####,
        ["True"]
    };

    raw_string_trailing_spaces_are_preserved => {
        r####"string text="""trail  """; Console.WriteLine(text.EndsWith("  "));"####,
        ["True"]
    };

    raw_string_backslash_is_literal_not_escape => {
        r####"string text="""C:\temp\file"""; Console.WriteLine(text.Contains(@"\"));"####,
        ["True"]
    };

    raw_string_tab_character_is_literal => {
        r####"string text="""a	b"""; Console.WriteLine(text.Length);"####,
        ["3"]
    };

    raw_interpolated_raw_string_embeds_variable => {
        r####"int count=7; string text=$"""items={count}"""; Console.WriteLine(text);"####,
        ["items=7"]
    };

    raw_interpolated_raw_string_embeds_expression => {
        r####"int a=2; int b=3; string text=$"""sum={a+b}"""; Console.WriteLine(text);"####,
        ["sum=5"]
    };

    raw_interpolated_raw_string_multiple_holes => {
        r####"string name="Ada"; int age=36; string text=$"""{name} is {age}"""; Console.WriteLine(text);"####,
        ["Ada is 36"]
    };

    raw_interpolated_raw_string_with_literal_braces => {
        r####"int n=1; string text=$"""value={n} end"""; Console.WriteLine(text.EndsWith(" end"));"####,
        ["True"]
    };

    raw_string_custom_delimiter_single_quote => {
        r####"string text=""""""quote "inside" here""""""; Console.WriteLine(text.Contains("inside"));"####,
        ["True"]
    };

    raw_string_custom_delimiter_allows_unescaped_quotes => {
        r####"string text=""""""a "b" c""""""; Console.WriteLine(text.Length>0);"####,
        ["True"]
    };

    raw_string_multiline_three_lines_counts_newlines => {
        r####"string text="""one
two
three"""; Console.WriteLine(text.Split('\n').Length);"####,
        ["3"]
    };

    raw_string_concatenation_with_regular_string => {
        r####"string text="""raw""" + "-suffix"; Console.WriteLine(text);"####,
        ["raw-suffix"]
    };

    raw_string_equality_compares_content => {
        r####"string a="""same"""; string b="""same"""; Console.WriteLine(a==b);"####,
        ["True"]
    };

    raw_string_inequality_detects_different_content => {
        r####"string a="""one"""; string b="""two"""; Console.WriteLine(a==b);"####,
        ["False"]
    };

    raw_string_indexer_reads_character => {
        r####"string text="""abcd"""; Console.WriteLine(text[2]);"####,
        ["99"]
    };

    raw_string_substring_extracts_suffix => {
        r####"string text="""abcdef"""; Console.WriteLine(text.Substring(4));"####,
        ["ef"]
    };

    raw_string_contains_search_finds_substring => {
        r####"string text="""hello world"""; Console.WriteLine(text.Contains("world"));"####,
        ["True"]
    };

    raw_string_replace_substitutes_text => {
        r####"string text="""foo-bar"""; Console.WriteLine(text.Replace("bar","baz"));"####,
        ["foo-baz"]
    };

    raw_string_to_upper_changes_case => {
        r####"string text="""abc"""; Console.WriteLine(text.ToUpper());"####,
        ["ABC"]
    };

    raw_string_trim_removes_whitespace_edges => {
        r####"string text="""  trim  """; Console.WriteLine(text.Trim());"####,
        ["trim"]
    };

    raw_string_split_on_comma => {
        r####"string text="""a,b,c"""; Console.WriteLine(text.Split(',')[1]);"####,
        ["b"]
    };

    raw_interpolated_multiline_preserves_newline_and_value => {
        r####"int id=9; string text=$"""id:
{id}"""; Console.WriteLine(text.Contains("9"));"####,
        ["True"]
    };

    raw_interpolated_with_format_specifier => {
        r####"double pi=3.14159; string text=$"""pi={pi:F2}"""; Console.WriteLine(text);"####,
        ["pi=3.14"]
    };

    raw_interpolated_with_alignment => {
        r####"int n=42; string text=$"""{n,5}"""; Console.WriteLine(text.Trim().Length>=2);"####,
        ["True"]
    };

    raw_string_starts_with_prefix => {
        r####"string text="""prefix-value"""; Console.WriteLine(text.StartsWith("prefix"));"####,
        ["True"]
    };

    raw_string_ends_with_suffix => {
        r####"string text="""value-suffix"""; Console.WriteLine(text.EndsWith("suffix"));"####,
        ["True"]
    };

    raw_string_insert_adds_middle_text => {
        r####"string text="""ac"""; Console.WriteLine(text.Insert(1,"b"));"####,
        ["abc"]
    };

    raw_string_remove_drops_characters => {
        r####"string text="""abcde"""; Console.WriteLine(text.Remove(1,2));"####,
        ["ade"]
    };

    raw_string_pad_left_adds_characters => {
        r####"string text="""7"""; Console.WriteLine(text.PadLeft(3,'0'));"####,
        ["007"]
    };

    raw_string_pad_right_adds_characters => {
        r####"string text="""7"""; Console.WriteLine(text.PadRight(3,'0'));"####,
        ["700"]
    };

    raw_string_compare_ordinal_ignore_case => {
        r####"string a="""Hello"""; string b="""hello"""; Console.WriteLine(string.Equals(a,b,System.StringComparison.OrdinalIgnoreCase));"####,
        ["True"]
    };

    raw_string_get_hash_code_consistent_for_same_content => {
        r####"string a="""hash"""; string b="""hash"""; Console.WriteLine(a.GetHashCode()==b.GetHashCode());"####,
        ["True"]
    };

    raw_string_is_null_or_empty_false_for_content => {
        r####"string text="""x"""; Console.WriteLine(string.IsNullOrEmpty(text));"####,
        ["False"]
    };

    raw_string_is_null_or_white_space_false_for_content => {
        r####"string text="""x"""; Console.WriteLine(string.IsNullOrWhiteSpace(text));"####,
        ["False"]
    };

    raw_string_join_with_separator => {
        r####"string text=string.Join("-",new string[]{"""a""","""b""","""c"""}); Console.WriteLine(text);"####,
        ["a-b-c"]
    };

    raw_interpolated_culture_invariant_numeric => {
        r####"double value=1234.5; string text=$"""{value.ToString(System.Globalization.CultureInfo.InvariantCulture)}"""; Console.WriteLine(text.Contains("."));"####,
        ["True"]
    };

    raw_string_line_count_in_multiline_literal => {
        r####"string text="""row1
row2
row3
row4"""; Console.WriteLine(text.Split('\n').Length);"####,
        ["4"]
    };

    raw_string_with_unicode_characters => {
        r####"string text="""café"""; Console.WriteLine(text.Contains("é"));"####,
        ["True"]
    };

    raw_interpolated_string_with_conditional_expression => {
        r####"int n=4; string text=$"""{ (n%2==0 ? "even" : "odd") }"""; Console.WriteLine(text.Trim());"####,
        ["even"]
    };

    raw_string_multiple_embedded_quotes_count => {
        r####"string text="""a "b" c "d" e"""; Console.WriteLine(text.Split('"').Length);"####,
        ["6"]
    };

    raw_string_literal_newline_not_escape_sequence => {
        r####"string text="""top
bottom"""; Console.WriteLine(text.IndexOf('\n')>0);"####,
        ["True"]
    };

    raw_interpolated_nested_string_field => {
        r####"class Item{public string Label="""tag""";} var item=new Item(); string text=$"""label={item.Label}"""; Console.WriteLine(text.Contains("tag"));"####,
        ["True"]
    };

    raw_string_repeat_doubles_content => {
        r####"string text="""ab"""; Console.WriteLine(text+text);"####,
        ["abab"]
    };
}
