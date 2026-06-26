//! LINQ quantifiers (`All`, `Any`, `Contains`, `SequenceEqual`) and `Chunk` partitioning.
//! GAP: boolean sequence tests and batching beyond existing `test_csharp_linq_*` coverage.

csharp_cases! {
    all_even_true_for_all_even => {
        r#"Console.WriteLine(new[]{2,4,6}.All(x=>x%2==0));"#,
        ["True"]
    };

    all_even_false_when_one_odd => {
        r#"Console.WriteLine(new[]{2,4,5}.All(x=>x%2==0));"#,
        ["False"]
    };

    all_empty_vacuous_true => {
        r#"Console.WriteLine(System.Array.Empty<int>().All(x=>x>100));"#,
        ["True"]
    };

    all_greater_than_zero_mixed => {
        r#"Console.WriteLine(new[]{1,2,0}.All(x=>x>0));"#,
        ["False"]
    };

    all_strings_non_empty => {
        r#"Console.WriteLine(new[]{"a","b"}.All(s=>s.Length>0));"#,
        ["True"]
    };

    all_strings_one_empty => {
        r#"Console.WriteLine(new[]{"a",""}.All(s=>s.Length>0));"#,
        ["False"]
    };

    any_with_predicate_true => {
        r#"Console.WriteLine(new[]{1,2,3}.Any(x=>x>2));"#,
        ["True"]
    };

    any_with_predicate_false => {
        r#"Console.WriteLine(new[]{1,2,3}.Any(x=>x>10));"#,
        ["False"]
    };

    any_without_predicate_nonempty => {
        r#"Console.WriteLine(new[]{1}.Any());"#,
        ["True"]
    };

    any_without_predicate_empty => {
        r#"Console.WriteLine(System.Array.Empty<int>().Any());"#,
        ["False"]
    };

    any_after_where_still_has_matches => {
        r#"Console.WriteLine(new[]{1,2,3,4}.Where(x=>x%2==0).Any());"#,
        ["True"]
    };

    any_after_where_no_matches => {
        r#"Console.WriteLine(new[]{1,3,5}.Where(x=>x%2==0).Any());"#,
        ["False"]
    };

    contains_present_value => {
        r#"Console.WriteLine(new[]{1,2,3}.Contains(2));"#,
        ["True"]
    };

    contains_absent_value => {
        r#"Console.WriteLine(new[]{1,2,3}.Contains(9));"#,
        ["False"]
    };

    contains_on_empty => {
        r#"Console.WriteLine(System.Array.Empty<int>().Contains(1));"#,
        ["False"]
    };

    contains_string_case_sensitive => {
        r#"Console.WriteLine(new[]{"A","b"}.Contains("a"));"#,
        ["False"]
    };

    contains_string_exact_match => {
        r#"Console.WriteLine(new[]{"A","b"}.Contains("A"));"#,
        ["True"]
    };

    contains_after_distinct => {
        r#"Console.WriteLine(new[]{1,1,2}.Distinct().Contains(2));"#,
        ["True"]
    };

    sequence_equal_identical_arrays => {
        r#"Console.WriteLine(new[]{1,2,3}.SequenceEqual(new[]{1,2,3}));"#,
        ["True"]
    };

    sequence_equal_different_lengths => {
        r#"Console.WriteLine(new[]{1,2,3}.SequenceEqual(new[]{1,2}));"#,
        ["False"]
    };

    sequence_equal_same_elements_different_order => {
        r#"Console.WriteLine(new[]{1,2,3}.SequenceEqual(new[]{3,2,1}));"#,
        ["False"]
    };

    sequence_equal_empty_arrays => {
        r#"Console.WriteLine(System.Array.Empty<int>().SequenceEqual(System.Array.Empty<int>()));"#,
        ["True"]
    };

    sequence_equal_one_empty => {
        r#"Console.WriteLine(new[]{1}.SequenceEqual(System.Array.Empty<int>()));"#,
        ["False"]
    };

    sequence_equal_strings => {
        r#"Console.WriteLine(new[]{"a","b"}.SequenceEqual(new[]{"a","b"}));"#,
        ["True"]
    };

    sequence_equal_after_orderby => {
        r#"Console.WriteLine(new[]{3,1,2}.OrderBy(x=>x).SequenceEqual(new[]{1,2,3}));"#,
        ["True"]
    };

    all_and_any_combined_count => {
        r#"var data=new[]{2,4,6,8};
Console.WriteLine(data.All(x=>x%2==0)?1:0);
Console.WriteLine(data.Any(x=>x>5)?1:0);"#,
        ["1", "1"]
    };

    contains_and_any_pipeline => {
        r#"var data=new[]{1,2,3,4};
Console.WriteLine(data.Contains(3)?1:0);
Console.WriteLine(data.Any(x=>x>3)?1:0);"#,
        ["1", "1"]
    };

    sequence_equal_self_via_skip_take => {
        r#"var a=new[]{1,2,3,4};
Console.WriteLine(a.Skip(1).Take(2).SequenceEqual(new[]{2,3}));"#,
        ["True"]
    };

    chunk_size_two_batch_count => {
        r#"Console.WriteLine(new[]{1,2,3,4,5}.Chunk(2).Count());"#,
        ["3"]
    };

    chunk_size_two_first_batch_length => {
        r#"Console.WriteLine(new[]{1,2,3,4,5}.Chunk(2).First().Length);"#,
        ["2"]
    };

    chunk_size_two_last_batch_length => {
        r#"Console.WriteLine(new[]{1,2,3,4,5}.Chunk(2).Last().Length);"#,
        ["1"]
    };

    chunk_size_three_batch_count => {
        r#"Console.WriteLine(new[]{1,2,3,4,5,6,7}.Chunk(3).Count());"#,
        ["3"]
    };

    chunk_size_larger_than_sequence_single_batch => {
        r#"Console.WriteLine(new[]{1,2}.Chunk(5).Count());"#,
        ["1"]
    };

    chunk_size_one_many_batches => {
        r#"Console.WriteLine(new[]{1,2,3}.Chunk(1).Count());"#,
        ["3"]
    };

    chunk_empty_sequence_count => {
        r#"Console.WriteLine(System.Array.Empty<int>().Chunk(2).Count());"#,
        ["0"]
    };

    chunk_sum_of_batch_lengths => {
        r#"var batches=new[]{1,2,3,4,5,6}.Chunk(2);
Console.WriteLine(batches.Sum(b=>b.Length));"#,
        ["6"]
    };

    chunk_first_batch_sum => {
        r#"Console.WriteLine(new[]{1,2,3,4}.Chunk(2).First().Sum());"#,
        ["3"]
    };

    chunk_last_batch_sum => {
        r#"Console.WriteLine(new[]{1,2,3,4,5}.Chunk(2).Last().Sum());"#,
        ["5"]
    };

    chunk_strings_batch_count => {
        r#"Console.WriteLine(new[]{"a","b","c","d","e"}.Chunk(2).Count());"#,
        ["3"]
    };

    partition_via_skip_take_page_count => {
        r#"var src=new[]{1,2,3,4,5,6};
int pageSize=2;
int pageCount=0;
for(int i=0;i<src.Length;i+=pageSize) pageCount++;
Console.WriteLine(pageCount);"#,
        ["3"]
    };

    partition_via_skip_take_second_page_sum => {
        r#"var src=new[]{1,2,3,4,5,6};
Console.WriteLine(src.Skip(2).Take(2).Sum());"#,
        ["7"]
    };

    partition_manual_window_count => {
        r#"var src=new[]{10,20,30,40,50};
int size=2;
int windows=0;
for(int i=0;i+size<=src.Length;i+=size) windows++;
Console.WriteLine(windows);"#,
        ["2"]
    };

    all_any_contains_combined_flags => {
        r#"var xs=new[]{1,2,3};
Console.WriteLine(xs.All(x=>x>0)?1:0);
Console.WriteLine(xs.Any(x=>x==2)?1:0);
Console.WriteLine(xs.Contains(4)?1:0);"#,
        ["1", "1", "0"]
    };

    sequence_equal_chunked_batches_same => {
        r#"var a=new[]{1,2,3,4};
var b=new[]{1,2,3,4};
Console.WriteLine(a.Chunk(2).SelectMany(x=>x).SequenceEqual(b));"#,
        ["True"]
    };

    chunk_then_all_batches_full_except_last => {
        r#"var batches=new[]{1,2,3,4,5,6,7}.Chunk(3);
Console.WriteLine(batches.Take(2).All(b=>b.Length==3)?1:0);
Console.WriteLine(batches.Last().Length);"#,
        ["1", "1"]
    };

    any_all_on_chunk_existence => {
        r#"var batches=new[]{1,2,3,4}.Chunk(2);
Console.WriteLine(batches.Any()?1:0);
Console.WriteLine(batches.All(b=>b.Length>0)?1:0);"#,
        ["1", "1"]
    };

    contains_in_chunk_flattened => {
        r#"var flat=new[]{1,2,3,4,5}.Chunk(2).SelectMany(x=>x);
Console.WriteLine(flat.Contains(5)?1:0);"#,
        ["1"]
    };

    sequence_equal_different_element_same_length => {
        r#"Console.WriteLine(new[]{1,2,3}.SequenceEqual(new[]{1,2,4}));"#,
        ["False"]
    };

    all_predicate_on_chunk_first_batch => {
        r#"Console.WriteLine(new[]{2,4,6,8}.Chunk(2).First().All(x=>x%2==0));"#,
        ["True"]
    };

    chunk_batch_count_via_select => {
        r#"Console.WriteLine(new[]{1,2,3,4,5,6,7,8}.Chunk(4).Select(b=>b.Length).Count());"#,
        ["2"]
    };
}
