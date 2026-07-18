//! UTF-8 string literals (`u8"..."`, C# 11) produce `ReadOnlySpan<byte>`.
//! GAP: no prior coverage of `u8` literal syntax or byte-span decoding.

csharp_cases! {
    utf8_literal_empty_has_zero_length => {
        r#"var bytes=u8""; Console.WriteLine(bytes.Length);"#,
        ["0"]
    };

    utf8_literal_ascii_length_matches_char_count => {
        r#"var bytes=u8"abc"; Console.WriteLine(bytes.Length);"#,
        ["3"]
    };

    utf8_literal_first_byte_is_uppercase_a => {
        r#"var bytes=u8"ABC"; Console.WriteLine(bytes[0]);"#,
        ["65"]
    };

    utf8_literal_second_byte_is_lowercase_b => {
        r#"var bytes=u8"abc"; Console.WriteLine(bytes[1]);"#,
        ["98"]
    };

    utf8_literal_third_byte_is_lowercase_c => {
        r#"var bytes=u8"abc"; Console.WriteLine(bytes[2]);"#,
        ["99"]
    };

    utf8_literal_decodes_to_string_via_encoding => {
        r#"var bytes=u8"hello"; Console.WriteLine(System.Text.Encoding.UTF8.GetString(bytes));"#,
        ["hello"]
    };

    utf8_literal_space_byte_value => {
        r#"var bytes=u8"a b"; Console.WriteLine(bytes[1]);"#,
        ["32"]
    };

    utf8_literal_digit_bytes => {
        r#"var bytes=u8"123"; Console.WriteLine(bytes[0]); Console.WriteLine(bytes[2]);"#,
        ["49", "51"]
    };

    utf8_literal_newline_escape_single_byte => {
        r#"var bytes=u8"a\nb"; Console.WriteLine(bytes[1]);"#,
        ["10"]
    };

    utf8_literal_tab_escape_single_byte => {
        r#"var bytes=u8"a\tb"; Console.WriteLine(bytes[1]);"#,
        ["9"]
    };

    utf8_literal_backslash_escape => {
        r#"var bytes=u8"a\\b"; Console.WriteLine(bytes[1]);"#,
        ["92"]
    };

    utf8_literal_quote_escape => {
        r#"var bytes=u8"a"b"; Console.WriteLine(bytes[1]);"#,
        ["34"]
    };

    utf8_literal_unicode_two_byte_sequence => {
        r#"var bytes=u8"é"; Console.WriteLine(bytes.Length);"#,
        ["2"]
    };

    utf8_literal_unicode_decodes_to_char => {
        r#"var bytes=u8"é"; Console.WriteLine(System.Text.Encoding.UTF8.GetString(bytes));"#,
        ["é"]
    };

    utf8_literal_cafe_unicode_length => {
        r#"var bytes=u8"café"; Console.WriteLine(bytes.Length);"#,
        ["5"]
    };

    utf8_literal_cafe_decodes_full_text => {
        r#"var bytes=u8"café"; Console.WriteLine(System.Text.Encoding.UTF8.GetString(bytes));"#,
        ["café"]
    };

    utf8_literal_single_byte_zero => {
        r#"var bytes=u8"\0"; Console.WriteLine(bytes[0]);"#,
        ["0"]
    };

    utf8_literal_hex_escape_byte => {
        r#"var bytes=u8"\x41"; Console.WriteLine(bytes[0]);"#,
        ["65"]
    };

    utf8_literal_mixed_case_ascii => {
        r#"var bytes=u8"AbC"; Console.WriteLine(bytes[1]);"#,
        ["98"]
    };

    utf8_literal_long_ascii_length => {
        r#"var bytes=u8"programming"; Console.WriteLine(bytes.Length);"#,
        ["11"]
    };

    utf8_literal_last_byte_of_hello => {
        r#"var bytes=u8"hello"; Console.WriteLine(bytes[4]);"#,
        ["111"]
    };

    utf8_literal_equals_manual_byte_array => {
        r#"var bytes=u8"hi"; Console.WriteLine(bytes[0]==104 && bytes[1]==105);"#,
        ["True"]
    };

    utf8_literal_slice_first_two_bytes => {
        r#"var bytes=u8"hello"; Console.WriteLine(bytes[0]); Console.WriteLine(bytes[1]);"#,
        ["104", "101"]
    };

    utf8_literal_copy_to_array_preserves_bytes => {
        r#"var bytes=u8"xy"; byte[] buf=new byte[2]; bytes.CopyTo(buf); Console.WriteLine(buf[0]); Console.WriteLine(buf[1]);"#,
        ["120", "121"]
    };

    utf8_literal_index_from_end => {
        r#"var bytes=u8"data"; Console.WriteLine(bytes[^1]);"#,
        ["97"]
    };

    utf8_literal_range_slice_length => {
        r#"var bytes=u8"abcdef"; var slice=bytes[2..5]; Console.WriteLine(slice.Length);"#,
        ["3"]
    };

    utf8_literal_range_slice_first_byte => {
        r#"var bytes=u8"abcdef"; var slice=bytes[2..5]; Console.WriteLine(slice[0]);"#,
        ["99"]
    };

    utf8_literal_sequence_equal_same_literal => {
        r#"var a=u8"same"; var b=u8"same"; Console.WriteLine(a.SequenceEqual(b));"#,
        ["True"]
    };

    utf8_literal_sequence_equal_different_literal => {
        r#"var a=u8"one"; var b=u8"two"; Console.WriteLine(a.SequenceEqual(b));"#,
        ["False"]
    };

    utf8_literal_starts_with_prefix_bytes => {
        r#"var bytes=u8"prefix"; Console.WriteLine(bytes.StartsWith(u8"pre"));"#,
        ["True"]
    };

    utf8_literal_ends_with_suffix_bytes => {
        r#"var bytes=u8"suffix"; Console.WriteLine(bytes.EndsWith(u8"fix"));"#,
        ["True"]
    };

    utf8_literal_index_of_byte_found => {
        r#"var bytes=u8"banana"; Console.WriteLine(bytes.IndexOf((byte)'a'));"#,
        ["1"]
    };

    utf8_literal_index_of_byte_not_found => {
        r#"var bytes=u8"banana"; Console.WriteLine(bytes.IndexOf((byte)'z'));"#,
        ["-1"]
    };

    utf8_literal_contains_byte => {
        r#"var bytes=u8"test"; Console.WriteLine(bytes.Contains((byte)'e'));"#,
        ["True"]
    };

    utf8_literal_foreach_sum_of_bytes => {
        r#"var bytes=u8"ab"; int sum=0; foreach(var b in bytes) sum+=b; Console.WriteLine(sum);"#,
        ["195"]
    };

    utf8_literal_backslash_in_content_is_literal => {
        r#"var bytes=u8"pa\\th"; Console.WriteLine(bytes[2]);"#,
        ["92"]
    };

    utf8_literal_concatenation_not_supported_use_separate => {
        r#"var left=u8"ab"; var right=u8"cd"; Console.WriteLine(left.Length+right.Length);"#,
        ["4"]
    };

    utf8_literal_empty_not_null_span => {
        r#"var bytes=u8""; Console.WriteLine(bytes.IsEmpty);"#,
        ["True"]
    };

    utf8_literal_non_empty_is_not_empty => {
        r#"var bytes=u8"x"; Console.WriteLine(bytes.IsEmpty);"#,
        ["False"]
    };

    utf8_literal_get_hash_code_consistent => {
        r#"var a=u8"hash"; var b=u8"hash"; Console.WriteLine(a.GetHashCode()==b.GetHashCode());"#,
        ["True"]
    };
}
