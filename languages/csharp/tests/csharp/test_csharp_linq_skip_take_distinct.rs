//! LINQ `Skip`, `Take`, `SkipWhile`, `TakeWhile`, `Distinct`, and `DistinctBy`.
//! GAP: paging/distinct-key operators beyond existing `test_csharp_linq_*` coverage.

csharp_cases! {
    skip_first_two_count_remaining => {
        r#"var r=new[]{10,20,30,40}.Skip(2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    skip_first_two_first_element => {
        r#"var r=new[]{10,20,30,40}.Skip(2);
Console.WriteLine(r.First());"#,
        ["30"]
    };

    skip_more_than_length_yields_empty_count => {
        r#"var r=new[]{1,2,3}.Skip(10);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    skip_zero_returns_full_count => {
        r#"var r=new[]{1,2,3}.Skip(0);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    skip_on_empty_sequence_count => {
        r#"var r=System.Array.Empty<int>().Skip(3);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    skip_one_then_take_two_count => {
        r#"var r=new[]{1,2,3,4,5}.Skip(1).Take(2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    skip_chained_twice_count => {
        r#"var r=new[]{1,2,3,4,5,6}.Skip(1).Skip(2);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    skip_last_element_via_skip_count => {
        r#"var r=new[]{1,2,3,4}.Skip(3);
Console.WriteLine(r.Count());"#,
        ["1"]
    };

    skip_strings_skip_prefix_count => {
        r#"var r=new[]{"a","b","c","d"}.Skip(2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    skip_after_orderby_count => {
        r#"var r=new[]{3,1,2}.OrderBy(x=>x).Skip(1);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    take_first_three_count => {
        r#"var r=new[]{10,20,30,40,50}.Take(3);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    take_first_three_last_taken => {
        r#"var r=new[]{10,20,30,40,50}.Take(3);
Console.WriteLine(r.Last());"#,
        ["30"]
    };

    take_more_than_length_returns_all_count => {
        r#"var r=new[]{1,2}.Take(10);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    take_zero_yields_empty_count => {
        r#"var r=new[]{1,2,3}.Take(0);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    take_on_empty_sequence_count => {
        r#"var r=System.Array.Empty<int>().Take(5);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    take_one_single_element => {
        r#"var r=new[]{99,100,101}.Take(1);
Console.WriteLine(r.Single());"#,
        ["99"]
    };

    take_after_skip_window_count => {
        r#"var r=new[]{1,2,3,4,5,6}.Skip(2).Take(2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    take_after_skip_window_sum => {
        r#"var r=new[]{1,2,3,4,5,6}.Skip(2).Take(2);
Console.WriteLine(r.Sum());"#,
        ["7"]
    };

    take_while_odd_prefix_count => {
        r#"var r=new[]{1,3,5,4,7}.TakeWhile(x=>x%2!=0);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    take_while_odd_prefix_sum => {
        r#"var r=new[]{1,3,5,4,7}.TakeWhile(x=>x%2!=0);
Console.WriteLine(r.Sum());"#,
        ["9"]
    };

    take_while_all_match_returns_full_count => {
        r#"var r=new[]{2,4,6}.TakeWhile(x=>x%2==0);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    take_while_none_match_yields_empty_count => {
        r#"var r=new[]{1,3,5}.TakeWhile(x=>x%2==0);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    take_while_strings_prefix_length_le_two_count => {
        r#"var r=new[]{"a","bb","ccc"}.TakeWhile(s=>s.Length<=2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    take_while_on_empty_count => {
        r#"var r=System.Array.Empty<int>().TakeWhile(x=>x<5);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    skip_while_less_than_three_count => {
        r#"var r=new[]{1,2,3,4,5}.SkipWhile(x=>x<3);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    skip_while_less_than_three_first => {
        r#"var r=new[]{1,2,3,4,5}.SkipWhile(x=>x<3);
Console.WriteLine(r.First());"#,
        ["3"]
    };

    skip_while_all_skipped_yields_empty_count => {
        r#"var r=new[]{1,2}.SkipWhile(x=>x<10);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    skip_while_none_skipped_returns_full_count => {
        r#"var r=new[]{5,6,7}.SkipWhile(x=>x<3);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    skip_while_strings_skip_short_prefix_count => {
        r#"var r=new[]{"a","bb","ccc","d"}.SkipWhile(s=>s.Length<3);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    skip_while_on_empty_count => {
        r#"var r=System.Array.Empty<int>().SkipWhile(x=>x<1);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    distinct_int_duplicates_count => {
        r#"var r=new[]{1,2,2,3,1,4}.Distinct();
Console.WriteLine(r.Count());"#,
        ["4"]
    };

    distinct_int_duplicates_sum => {
        r#"var r=new[]{1,2,2,3,1,4}.Distinct();
Console.WriteLine(r.Sum());"#,
        ["10"]
    };

    distinct_all_same_count => {
        r#"var r=new[]{7,7,7,7}.Distinct();
Console.WriteLine(r.Count());"#,
        ["1"]
    };

    distinct_empty_count => {
        r#"var r=System.Array.Empty<int>().Distinct();
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    distinct_strings_case_sensitive_count => {
        r#"var r=new[]{"A","a","A","b"}.Distinct();
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    distinct_after_orderby_count => {
        r#"var r=new[]{3,1,2,1}.OrderBy(x=>x).Distinct();
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    distinct_preserves_first_occurrence_foreach => {
        r#"var r=new[]{2,1,2,3,1}.Distinct();
foreach(var n in r) Console.WriteLine(n);"#,
        ["2", "1", "3"]
    };

    distinct_by_length_count => {
        r#"var r=new[]{"a","bb","c","dd","eee"}.DistinctBy(s=>s.Length);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    distinct_by_length_first_of_each_group => {
        r#"var r=new[]{"a","bb","c","dd","eee"}.DistinctBy(s=>s.Length);
foreach(var s in r) Console.WriteLine(s);"#,
        ["a", "bb", "eee"]
    };

    distinct_by_first_char_count => {
        r#"var r=new[]{"cat","car","dog","dot"}.DistinctBy(s=>s[0]);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    distinct_by_first_char_first_elements => {
        r#"var r=new[]{"cat","car","dog","dot"}.DistinctBy(s=>s[0]);
Console.WriteLine(r.First()); Console.WriteLine(r.Last());"#,
        ["cat", "dog"]
    };

    distinct_by_mod_two_count => {
        r#"var r=new[]{1,2,3,4,5,6}.DistinctBy(n=>n%2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    distinct_by_mod_two_sum => {
        r#"var r=new[]{1,2,3,4,5,6}.DistinctBy(n=>n%2);
Console.WriteLine(r.Sum());"#,
        ["3"]
    };

    distinct_by_on_empty_count => {
        r#"var r=System.Array.Empty<int>().DistinctBy(n=>n);
Console.WriteLine(r.Count());"#,
        ["0"]
    };

    distinct_by_all_same_key_count => {
        r#"var r=new[]{10,20,30}.DistinctBy(n=>0);
Console.WriteLine(r.Count());"#,
        ["1"]
    };

    distinct_by_record_key_count => {
        r#"var r=new[]{(K:1,V:"a"),(K:1,V:"b"),(K:2,V:"c")}.DistinctBy(t=>t.K);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    distinct_by_record_key_first_values => {
        r#"var r=new[]{(K:1,V:"a"),(K:1,V:"b"),(K:2,V:"c")}.DistinctBy(t=>t.K);
Console.WriteLine(r.First().V); Console.WriteLine(r.Last().V);"#,
        ["a", "c"]
    };

    skip_take_distinct_pipeline_count => {
        r#"var r=new[]{1,2,2,3,3,4,5}.Skip(1).Take(5).Distinct();
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    skip_while_take_while_window_count => {
        r#"var r=new[]{1,2,3,4,5,6,7}.SkipWhile(x=>x<3).TakeWhile(x=>x<6);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    skip_while_take_while_window_sum => {
        r#"var r=new[]{1,2,3,4,5,6,7}.SkipWhile(x=>x<3).TakeWhile(x=>x<6);
Console.WriteLine(r.Sum());"#,
        ["12"]
    };

    distinct_by_then_order_count => {
        r#"var r=new[]{"zzz","a","bb","c","dd"}.DistinctBy(s=>s.Length).OrderBy(s=>s);
Console.WriteLine(r.Count());"#,
        ["3"]
    };

    take_while_then_skip_count => {
        r#"var r=new[]{1,2,3,4,5}.TakeWhile(x=>x<5).Skip(1);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    skip_then_distinct_by_count => {
        r#"var r=new[]{1,1,2,2,3,3}.Skip(2).DistinctBy(n=>n);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    distinct_then_take_count => {
        r#"var r=new[]{5,1,5,2,3,2}.Distinct().Take(2);
Console.WriteLine(r.Count());"#,
        ["2"]
    };

    paging_skip_take_repeat_page_two_count => {
        r#"var src=new[]{1,2,3,4,5,6,7,8,9};
var page=src.Skip(3).Take(3);
Console.WriteLine(page.Count());"#,
        ["3"]
    };

    paging_skip_take_repeat_page_two_sum => {
        r#"var src=new[]{1,2,3,4,5,6,7,8,9};
var page=src.Skip(3).Take(3);
Console.WriteLine(page.Sum());"#,
        ["12"]
    };
}
