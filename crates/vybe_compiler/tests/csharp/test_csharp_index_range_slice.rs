//! Index-from-end (`^n`), `Range` (`..`), and slicing on arrays and strings.
//! GAP: deep slice/index coverage beyond basic smoke tests in `ranges_indices`.

use crate::csharp_cases;

csharp_cases! {
    index_from_end_two_reads_second_to_last_array_element => {
        r#"int[] data={10,20,30,40}; Console.WriteLine(data[^2]);"#,
        ["30"]
    };

    index_from_end_three_reads_third_from_end => {
        r#"int[] data={1,2,3,4,5,6}; Console.WriteLine(data[^3]);"#,
        ["4"]
    };

    index_from_end_on_single_element_array => {
        r#"int[] data={99}; Console.WriteLine(data[^1]);"#,
        ["99"]
    };

    index_from_end_zero_is_past_end_sentinel_in_range => {
        r#"int[] data={1,2,3}; var slice=data[1..^0]; Console.WriteLine(slice.Length);"#,
        ["2"]
    };

    range_closed_start_open_end_excludes_last_element => {
        r#"int[] data={5,6,7,8,9}; var slice=data[1..^1]; Console.WriteLine(slice.Length); Console.WriteLine(slice[0]); Console.WriteLine(slice[1]);"#,
        ["3", "6", "7"]
    };

    range_from_index_from_end_start_to_open_end => {
        r#"int[] data={1,2,3,4,5}; var slice=data[^3..]; Console.WriteLine(slice.Length); Console.WriteLine(slice[0]);"#,
        ["3", "3"]
    };

    range_open_start_to_index_from_end_end => {
        r#"int[] data={1,2,3,4,5}; var slice=data[..^2]; Console.WriteLine(slice.Length); Console.WriteLine(slice[2]);"#,
        ["3", "3"]
    };

    range_both_indices_from_end => {
        r#"int[] data={10,20,30,40,50}; var slice=data[^4..^1]; Console.WriteLine(slice.Length); Console.WriteLine(slice[0]); Console.WriteLine(slice[2]);"#,
        ["3", "20", "40"]
    };

    range_single_element_slice_length_one => {
        r#"int[] data={9,8,7}; var slice=data[1..2]; Console.WriteLine(slice.Length); Console.WriteLine(slice[0]);"#,
        ["1", "8"]
    };

    range_empty_slice_when_start_equals_end => {
        r#"int[] data={1,2,3}; var slice=data[2..2]; Console.WriteLine(slice.Length);"#,
        ["0"]
    };

    range_full_array_via_explicit_bounds => {
        r#"int[] data={4,5,6}; var slice=data[0..3]; Console.WriteLine(slice.Length); Console.WriteLine(slice[2]);"#,
        ["3", "6"]
    };

    range_open_end_from_zero_returns_copy_of_all_elements => {
        r#"int[] data={11,22,33}; var slice=data[0..]; Console.WriteLine(slice.Length);"#,
        ["3"]
    };

    range_open_start_to_length_returns_prefix => {
        r#"int[] data={1,2,3,4}; var slice=data[..2]; Console.WriteLine(slice[0]); Console.WriteLine(slice[1]);"#,
        ["1", "2"]
    };

    range_on_char_array_produces_char_slice => {
        r#"char[] letters={'a','b','c','d'}; var slice=letters[1..3]; Console.WriteLine(slice.Length); Console.WriteLine(slice[0]);"#,
        ["2", "98"]
    };

    string_range_middle_segment => {
        r#"string text="abcdef"; Console.WriteLine(text[2..5]);"#,
        ["cde"]
    };

    string_range_open_start_prefix => {
        r#"string text="hello"; Console.WriteLine(text[..2]);"#,
        ["he"]
    };

    string_range_open_end_suffix => {
        r#"string text="hello"; Console.WriteLine(text[3..]);"#,
        ["lo"]
    };

    string_range_from_end_indices => {
        r#"string text="program"; Console.WriteLine(text[^4..^1]);"#,
        ["gra"]
    };

    string_range_full_copy_via_open_bounds => {
        r#"string text="abc"; Console.WriteLine(text[..]);"#,
        ["abc"]
    };

    string_index_from_end_reads_last_char_code => {
        r#"string text="xy"; Console.WriteLine(text[^1]);"#,
        ["121"]
    };

    string_single_char_range => {
        r#"string text="dart"; Console.WriteLine(text[1..2]);"#,
        ["a"]
    };

    array_slice_first_element_preserved => {
        r#"int[] data={100,200,300}; var slice=data[..1]; Console.WriteLine(slice[0]);"#,
        ["100"]
    };

    array_slice_last_element_via_from_end => {
        r#"int[] data={100,200,300}; var slice=data[^1..]; Console.WriteLine(slice[0]);"#,
        ["300"]
    };

    array_slice_tail_two_elements => {
        r#"int[] data={1,2,3,4,5}; var slice=data[3..]; Console.WriteLine(slice.Length); Console.WriteLine(slice[1]);"#,
        ["2", "5"]
    };

    array_slice_interior_skips_both_ends => {
        r#"int[] data={0,1,2,3,4}; var slice=data[1..4]; Console.WriteLine(slice[0]); Console.WriteLine(slice[2]);"#,
        ["1", "3"]
    };

    index_variable_from_end_used_in_access => {
        r#"int[] data={5,10,15}; System.Index idx=^2; Console.WriteLine(data[idx]);"#,
        ["10"]
    };

    range_variable_used_for_array_slice => {
        r#"int[] data={2,4,6,8}; System.Range r=new System.Range(1,3); var slice=data[r]; Console.WriteLine(slice.Length); Console.WriteLine(slice[1]);"#,
        ["2", "6"]
    };

    range_start_from_end_end_open => {
        r#"int[] data={1,2,3,4}; System.Range r=^2..; var slice=data[r]; Console.WriteLine(slice[0]); Console.WriteLine(slice[1]);"#,
        ["3", "4"]
    };

    range_end_from_end_start_open => {
        r#"int[] data={1,2,3,4}; System.Range r=..^1; var slice=data[r]; Console.WriteLine(slice.Length);"#,
        ["3"]
    };

    string_empty_range_produces_empty_substring => {
        r#"string text="abc"; Console.WriteLine(text[1..1].Length);"#,
        ["0"]
    };

    array_slice_assign_is_independent_copy => {
        r#"int[] data={1,2,3}; var slice=data[0..2]; slice[0]=9; Console.WriteLine(data[0]); Console.WriteLine(slice[0]);"#,
        ["1", "9"]
    };

    string_slice_does_not_mutate_original => {
        r#"string text="keep"; var part=text[1..3]; Console.WriteLine(text); Console.WriteLine(part);"#,
        ["keep", "ee"]
    };

    range_on_long_array_large_offset => {
        r#"int[] data={0,1,2,3,4,5,6,7,8,9}; var slice=data[7..]; Console.WriteLine(slice[0]); Console.WriteLine(slice[2]);"#,
        ["7", "9"]
    };

    index_from_end_on_two_element_array => {
        r#"int[] data={7,8}; Console.WriteLine(data[^1]); Console.WriteLine(data[^2]);"#,
        ["8", "7"]
    };

    string_range_unicode_characters => {
        r#"string text="café"; Console.WriteLine(text[1..3]);"#,
        ["af"]
    };

    array_range_zero_length_at_start => {
        r#"int[] data={1,2,3}; var slice=data[0..0]; Console.WriteLine(slice.Length);"#,
        ["0"]
    };

    array_range_zero_length_at_end => {
        r#"int[] data={1,2,3}; var slice=data[3..3]; Console.WriteLine(slice.Length);"#,
        ["0"]
    };

    string_range_to_end_from_second_char => {
        r#"string text="testing"; Console.WriteLine(text[1..]);"#,
        ["esting"]
    };

    array_slice_all_but_first => {
        r#"int[] data={9,8,7,6}; var slice=data[1..]; Console.WriteLine(string.Join(",",slice));"#,
        ["8,7,6"]
    };

    array_slice_all_but_last => {
        r#"int[] data={9,8,7,6}; var slice=data[..^1]; Console.WriteLine(string.Join(",",slice));"#,
        ["9,8,7"]
    };

    string_index_from_end_two => {
        r#"string text="abcde"; Console.WriteLine(text[^2]);"#,
        ["100"]
    };

    array_index_from_end_four_on_six_elements => {
        r#"int[] data={10,20,30,40,50,60}; Console.WriteLine(data[^4]);"#,
        ["30"]
    };

    range_half_open_slice_length_computed => {
        r#"int[] data={1,2,3,4,5,6,7}; var slice=data[2..5]; Console.WriteLine(slice.Length);"#,
        ["3"]
    };

    string_range_single_char_at_start => {
        r#"string text="open"; Console.WriteLine(text[0..1]);"#,
        ["o"]
    };

    string_range_single_char_at_end => {
        r#"string text="open"; Console.WriteLine(text[3..4]);"#,
        ["n"]
    };

    array_range_from_end_spanning_three => {
        r#"int[] data={1,2,3,4,5}; var slice=data[^5..^2]; Console.WriteLine(slice.Length); Console.WriteLine(slice[2]);"#,
        ["3", "3"]
    };

    index_from_end_on_string_with_spaces => {
        r#"string text="a b c"; Console.WriteLine(text[^2]);"#,
        ["32"]
    };

    array_slice_second_half_length => {
        r#"int[] data={1,2,3,4}; var slice=data[2..4]; Console.WriteLine(slice[0]); Console.WriteLine(slice[1]);"#,
        ["3", "4"]
    };

    string_range_between_spaces => {
        r#"string text="x y z"; Console.WriteLine(text[2..4]);"#,
        ["y "]
    };

    array_range_open_end_from_middle_index => {
        r#"int[] data={5,10,15,20,25}; var slice=data[2..]; Console.WriteLine(slice.Length); Console.WriteLine(slice[0]);"#,
        ["3", "15"]
    };

    index_from_end_one_matches_length_minus_one => {
        r#"int[] data={3,6,9}; Console.WriteLine(data[data.Length-1]); Console.WriteLine(data[^1]);"#,
        ["9", "9"]
    };

    string_full_range_equals_original => {
        r#"string text="same"; Console.WriteLine(text[..]==text);"#,
        ["True"]
    };

    array_slice_foreach_preserves_order => {
        r#"int[] data={1,2,3,4}; var slice=data[1..3]; int sum=0; foreach(var n in slice) sum+=n; Console.WriteLine(sum);"#,
        ["5"]
    };

    range_on_byte_array_slice => {
        r#"byte[] data={10,20,30,40}; var slice=data[1..3]; Console.WriteLine(slice[0]); Console.WriteLine(slice[1]);"#,
        ["20", "30"]
    };

    string_range_after_first_word => {
        r#"string text="hello world"; Console.WriteLine(text[6..11]);"#,
        ["world"]
    };

    array_range_one_before_end => {
        r#"int[] data={2,4,6,8,10}; var slice=data[..^1]; Console.WriteLine(slice[slice.Length-1]);"#,
        ["8"]
    };

    index_from_end_on_empty_range_start_marker => {
        r#"int[] data={1,2}; var slice=data[2..]; Console.WriteLine(slice.Length);"#,
        ["0"]
    };
}
