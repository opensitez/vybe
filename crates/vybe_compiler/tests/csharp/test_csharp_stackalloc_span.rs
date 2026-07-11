//! `stackalloc`, `Span<T>`, slicing, and stackalloc initializers.
//! GAP: memory/span coverage is thin in the existing suite.

csharp_cases! {
    stackalloc_int_zero_length_span_has_zero_length => {
        r#"System.Span<int> buf=stackalloc int[0]; Console.WriteLine(buf.Length);"#,
        ["0"]
    };

    stackalloc_int_single_element_initializer => {
        r#"System.Span<int> buf=stackalloc int[1]{42}; Console.WriteLine(buf[0]);"#,
        ["42"]
    };

    stackalloc_int_three_element_initializer_reads_middle => {
        r#"System.Span<int> buf=stackalloc int[3]{10,20,30}; Console.WriteLine(buf[1]);"#,
        ["20"]
    };

    stackalloc_int_initializer_sets_last_element => {
        r#"System.Span<int> buf=stackalloc int[4]{1,2,3,4}; Console.WriteLine(buf[3]);"#,
        ["4"]
    };

    stackalloc_int_write_through_span_persists_in_buffer => {
        r#"System.Span<int> buf=stackalloc int[2]{1,2}; buf[1]=99; Console.WriteLine(buf[1]);"#,
        ["99"]
    };

    stackalloc_byte_buffer_stores_ascii_code => {
        r#"System.Span<byte> buf=stackalloc byte[2]{65,66}; Console.WriteLine(buf[0]);"#,
        ["65"]
    };

    stackalloc_char_buffer_reads_first_character_code => {
        r#"System.Span<char> buf=stackalloc char[3]{'a','b','c'}; Console.WriteLine(buf[0]);"#,
        ["97"]
    };

    stackalloc_double_buffer_reads_fractional_value => {
        r#"System.Span<double> buf=stackalloc double[2]{1.5,2.5}; Console.WriteLine(buf[1]);"#,
        ["2.5"]
    };

    stackalloc_without_unsafe_in_safe_context => {
        r#"System.Span<int> nums=stackalloc int[3]{7,8,9}; Console.WriteLine(nums[2]);"#,
        ["9"]
    };

    span_from_stackalloc_reports_correct_length => {
        r#"System.Span<int> span=stackalloc int[5]{1,2,3,4,5}; Console.WriteLine(span.Length);"#,
        ["5"]
    };

    span_slice_start_one_reduces_length_by_one => {
        r#"System.Span<int> span=stackalloc int[4]{1,2,3,4}; var tail=span.Slice(1); Console.WriteLine(tail.Length);"#,
        ["3"]
    };

    span_slice_start_and_length_reads_subrange => {
        r#"System.Span<int> span=stackalloc int[5]{10,20,30,40,50}; var mid=span.Slice(1,2); Console.WriteLine(mid[0]); Console.WriteLine(mid[1]);"#,
        ["20", "30"]
    };

    span_slice_to_end_reads_last_element => {
        r#"System.Span<int> span=stackalloc int[3]{5,6,7}; var last=span.Slice(2); Console.WriteLine(last[0]);"#,
        ["7"]
    };

    readonly_span_from_stackalloc_is_read_only_view => {
        r#"System.ReadOnlySpan<int> view=stackalloc int[2]{3,4}; Console.WriteLine(view[1]);"#,
        ["4"]
    };

    stackalloc_span_fill_sets_all_elements => {
        r#"System.Span<int> span=stackalloc int[3]; span.Fill(9); Console.WriteLine(span[0]); Console.WriteLine(span[2]);"#,
        ["9", "9"]
    };

    stackalloc_span_clear_zeroes_elements => {
        r#"System.Span<int> span=stackalloc int[2]{5,6}; span.Clear(); Console.WriteLine(span[0]); Console.WriteLine(span[1]);"#,
        ["0", "0"]
    };

    stackalloc_span_copy_to_copies_values => {
        r#"System.Span<int> src=stackalloc int[2]{11,22}; System.Span<int> dst=stackalloc int[2]; src.CopyTo(dst); Console.WriteLine(dst[1]);"#,
        ["22"]
    };

    stackalloc_span_index_from_end_reads_first_slot => {
        r#"System.Span<int> span=stackalloc int[3]{8,9,10}; Console.WriteLine(span[^3]);"#,
        ["8"]
    };

    stackalloc_span_index_from_end_reads_last_slot => {
        r#"System.Span<int> span=stackalloc int[3]{8,9,10}; Console.WriteLine(span[^1]);"#,
        ["10"]
    };

    stackalloc_span_is_empty_false_for_non_zero_length => {
        r#"System.Span<int> span=stackalloc int[1]{1}; Console.WriteLine(span.IsEmpty);"#,
        ["False"]
    };

    stackalloc_span_is_empty_true_for_zero_length => {
        r#"System.Span<int> span=stackalloc int[0]; Console.WriteLine(span.IsEmpty);"#,
        ["True"]
    };

    stackalloc_span_try_copy_to_succeeds_when_dest_large_enough => {
        r#"System.Span<int> src=stackalloc int[2]{3,4}; System.Span<int> dst=stackalloc int[3]; Console.WriteLine(src.TryCopyTo(dst));"#,
        ["True"]
    };

    stackalloc_span_try_copy_to_fails_when_dest_too_small => {
        r#"System.Span<int> src=stackalloc int[3]{1,2,3}; System.Span<int> dst=stackalloc int[2]; Console.WriteLine(src.TryCopyTo(dst));"#,
        ["False"]
    };

    memory_wraps_stackalloc_backing_buffer_via_constructor => {
        r#"System.Memory<int> mem=new System.Memory<int>(stackalloc int[2]{1,2}); Console.WriteLine(mem.Length);"#,
        ["2"]
    };

    memory_span_reads_element_from_stackalloc_backing => {
        r#"System.Memory<int> mem=new System.Memory<int>(stackalloc int[3]{4,5,6}); Console.WriteLine(mem.Span[1]);"#,
        ["5"]
    };

    memory_span_write_updates_underlying_stackalloc_buffer => {
        r#"System.Memory<int> mem=new System.Memory<int>(stackalloc int[2]{1,2}); mem.Span[0]=77; Console.WriteLine(mem.Span[0]);"#,
        ["77"]
    };

    stackalloc_span_overwrite_first_element_after_init => {
        r#"System.Span<int> span=stackalloc int[3]{1,2,3}; span[0]=100; Console.WriteLine(span[0]);"#,
        ["100"]
    };

    stackalloc_span_iterate_sum_via_index_loop => {
        r#"System.Span<int> span=stackalloc int[3]{1,2,3}; int sum=0; for(int i=0;i<span.Length;i++){sum+=span[i];} Console.WriteLine(sum);"#,
        ["6"]
    };

    stackalloc_long_buffer_reads_expected_value => {
        r#"System.Span<long> span=stackalloc long[2]{10000000000L,20000000000L}; Console.WriteLine(span[0]>0);"#,
        ["True"]
    };

    stackalloc_bool_buffer_stores_true_literal => {
        r#"System.Span<bool> span=stackalloc bool[2]{true,false}; Console.WriteLine(span[0]);"#,
        ["True"]
    };

    stackalloc_float_buffer_reads_single_precision_value => {
        r#"System.Span<float> span=stackalloc float[2]{1.25f,2.5f}; Console.WriteLine(span[0]==1.25f);"#,
        ["True"]
    };

    stackalloc_span_slice_empty_at_end_has_zero_length => {
        r#"System.Span<int> span=stackalloc int[2]{1,2}; var empty=span.Slice(2); Console.WriteLine(empty.Length);"#,
        ["0"]
    };

    stackalloc_span_two_slices_compose_expected_element => {
        r#"System.Span<int> span=stackalloc int[6]{1,2,3,4,5,6}; var inner=span.Slice(2,2); Console.WriteLine(inner[1]);"#,
        ["4"]
    };

    stackalloc_span_starts_with_matching_prefix => {
        r#"System.Span<int> span=stackalloc int[3]{1,2,3}; System.ReadOnlySpan<int> prefix=stackalloc int[2]{1,2}; Console.WriteLine(span.StartsWith(prefix));"#,
        ["True"]
    };

    stackalloc_span_sequence_equal_compares_elementwise => {
        r#"System.Span<int> a=stackalloc int[2]{7,8}; System.Span<int> b=stackalloc int[2]{7,8}; Console.WriteLine(a.SequenceEqual(b));"#,
        ["True"]
    };

    stackalloc_span_sequence_equal_detects_mismatch => {
        r#"System.Span<int> a=stackalloc int[2]{7,8}; System.Span<int> b=stackalloc int[2]{7,9}; Console.WriteLine(a.SequenceEqual(b));"#,
        ["False"]
    };

    stackalloc_span_to_array_materializes_values => {
        r#"System.Span<int> span=stackalloc int[2]{12,34}; int[] arr=span.ToArray(); Console.WriteLine(arr[1]);"#,
        ["34"]
    };

    stackalloc_span_from_array_as_span_then_slice => {
        r#"int[] data={1,2,3,4}; System.Span<int> span=data.AsSpan(1,2); Console.WriteLine(span[0]); Console.WriteLine(span[1]);"#,
        ["2", "3"]
    };

    stackalloc_span_from_string_as_span_reads_char => {
        r#"System.ReadOnlySpan<char> chars="abcd".AsSpan(1,2); Console.WriteLine(chars[0]); Console.WriteLine(chars[1]);"#,
        ["98", "99"]
    };

    stackalloc_span_overlapping_copy_to_self_is_allowed => {
        r#"System.Span<int> span=stackalloc int[3]{1,2,3}; span.CopyTo(span.Slice(1)); Console.WriteLine(span[2]);"#,
        ["2"]
    };

    stackalloc_span_reverse_in_place_changes_order => {
        r#"System.Span<int> span=stackalloc int[3]{1,2,3}; span.Reverse(); Console.WriteLine(span[0]); Console.WriteLine(span[2]);"#,
        ["3", "1"]
    };

    stackalloc_span_binary_search_finds_existing_value => {
        r#"System.Span<int> span=stackalloc int[5]{1,3,5,7,9}; Console.WriteLine(System.MemoryExtensions.BinarySearch(span,5));"#,
        ["2"]
    };

    stackalloc_span_index_of_finds_element_offset => {
        r#"System.Span<int> span=stackalloc int[4]{10,20,30,40}; Console.WriteLine(span.IndexOf(30));"#,
        ["2"]
    };

    stackalloc_span_last_index_of_finds_last_match => {
        r#"System.Span<int> span=stackalloc int[4]{1,2,2,3}; Console.WriteLine(span.LastIndexOf(2));"#,
        ["2"]
    };

    stackalloc_span_contains_reports_present_value => {
        r#"System.Span<int> span=stackalloc int[3]{4,5,6}; Console.WriteLine(span.Contains(5));"#,
        ["True"]
    };

    stackalloc_span_contains_reports_missing_value => {
        r#"System.Span<int> span=stackalloc int[3]{4,5,6}; Console.WriteLine(span.Contains(99));"#,
        ["False"]
    };

    stackalloc_span_trim_start_skips_leading_match => {
        r#"System.Span<int> span=stackalloc int[4]{0,0,1,2}; var trimmed=span.TrimStart(0); Console.WriteLine(trimmed[0]);"#,
        ["1"]
    };

    stackalloc_span_trim_end_skips_trailing_match => {
        r#"System.Span<int> span=stackalloc int[4]{1,2,0,0}; var trimmed=span.TrimEnd(0); Console.WriteLine(trimmed[^1]);"#,
        ["2"]
    };

    stackalloc_span_mismatch_reports_first_difference_index => {
        r#"System.ReadOnlySpan<int> a=stackalloc int[3]{1,2,3}; System.ReadOnlySpan<int> b=stackalloc int[3]{1,9,3}; Console.WriteLine(a.Mismatch(b));"#,
        ["1"]
    };

    stackalloc_span_copy_to_existing_array_writes_values => {
        r#"System.Span<int> src=stackalloc int[2]{8,9}; int[] dst=new int[2]; src.CopyTo(dst); Console.WriteLine(dst[1]);"#,
        ["9"]
    };

    stackalloc_span_enumerator_moves_through_elements => {
        r#"System.Span<int> span=stackalloc int[3]{2,4,6}; int sum=0; foreach(int v in span){sum+=v;} Console.WriteLine(sum);"#,
        ["12"]
    };

    stackalloc_span_slice_preserves_element_after_mutation => {
        r#"System.Span<int> span=stackalloc int[4]{1,2,3,4}; var part=span.Slice(1,2); part[0]=88; Console.WriteLine(span[1]);"#,
        ["88"]
    };

    stackalloc_span_length_after_slice_one_from_three => {
        r#"System.Span<int> span=stackalloc int[3]{5,6,7}; Console.WriteLine(span.Slice(1).Length);"#,
        ["2"]
    };

    stackalloc_span_default_struct_has_zero_length => {
        r#"System.Span<int> span=default; Console.WriteLine(span.Length);"#,
        ["0"]
    };

    stackalloc_span_empty_property_matches_zero_length => {
        r#"System.Span<int> span=stackalloc int[0]; Console.WriteLine(System.Span<int>.Empty.Length);"#,
        ["0"]
    };
}
