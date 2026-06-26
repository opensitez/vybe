//! LINQ `Aggregate`, `MinBy`, `MaxBy`, `ElementAt`, `Single`, and `SingleOrDefault`.
//! GAP: seed folding and element pickers beyond existing `test_csharp_linq_*` coverage.

csharp_cases! {
    aggregate_seed_sum_integers => {
        r#"Console.WriteLine(new[]{1,2,3,4}.Aggregate(0,(acc,x)=>acc+x));"#,
        ["10"]
    };

    aggregate_seed_product_long => {
        r#"Console.WriteLine(new[]{1,2,3,4}.Aggregate(1L,(acc,n)=>acc*n));"#,
        ["24"]
    };

    aggregate_seed_string_concat_length => {
        r#"var s=new[]{"a","b","c"}.Aggregate("",(acc,x)=>acc+x);
Console.WriteLine(s.Length);"#,
        ["3"]
    };

    aggregate_seed_string_concat_value => {
        r#"var s=new[]{"a","b","c"}.Aggregate("",(acc,x)=>acc+x);
Console.WriteLine(s);"#,
        ["abc"]
    };

    aggregate_seed_max_running => {
        r#"var max=new[]{3,1,4,1,5}.Aggregate(int.MinValue,(acc,x)=>x>acc?x:acc);
Console.WriteLine(max);"#,
        ["5"]
    };

    aggregate_seed_min_running => {
        r#"var min=new[]{3,1,4,1,5}.Aggregate(int.MaxValue,(acc,x)=>x<acc?x:acc);
Console.WriteLine(min);"#,
        ["1"]
    };

    aggregate_seed_count_via_fold => {
        r#"var count=new[]{1,2,3}.Aggregate(0,(acc,x)=>acc+1);
Console.WriteLine(count);"#,
        ["3"]
    };

    aggregate_seed_empty_returns_seed => {
        r#"Console.WriteLine(System.Array.Empty<int>().Aggregate(99,(acc,x)=>acc+x));"#,
        ["99"]
    };

    aggregate_no_seed_sum_via_lambda => {
        r#"Console.WriteLine(new[]{1,2,3}.Aggregate((a,b)=>a+b));"#,
        ["6"]
    };

    aggregate_no_seed_product => {
        r#"Console.WriteLine(new[]{2,3,4}.Aggregate((a,b)=>a*b));"#,
        ["24"]
    };

    aggregate_no_seed_single_element => {
        r#"Console.WriteLine(new[]{42}.Aggregate((a,b)=>a+b));"#,
        ["42"]
    };

    aggregate_seed_with_index_sum => {
        r#"var sum=new[]{10,20,30}.Aggregate(0,(acc,x,i)=>acc+x+i);
Console.WriteLine(sum);"#,
        ["63"]
    };

    aggregate_seed_build_comma_list_length => {
        r#"var text=new[]{1,2,3}.Aggregate("",(acc,x)=>acc==""?x.ToString():acc+","+x);
Console.WriteLine(text.Length);"#,
        ["5"]
    };

    aggregate_seed_build_comma_list_value => {
        r#"var text=new[]{1,2,3}.Aggregate("",(acc,x)=>acc==""?x.ToString():acc+","+x);
Console.WriteLine(text);"#,
        ["1,2,3"]
    };

    min_by_shortest_word => {
        r#"Console.WriteLine(new[]{"aa","b","ccc"}.MinBy(w=>w.Length));"#,
        ["b"]
    };

    min_by_shortest_word_length => {
        r#"Console.WriteLine(new[]{"aa","b","ccc"}.MinBy(w=>w.Length).Length);"#,
        ["1"]
    };

    min_by_largest_number_by_abs => {
        r#"Console.WriteLine(new[]{-5,2,-1}.MinBy(n=>System.Math.Abs(n)));"#,
        ["-1"]
    };

    min_by_first_of_tie => {
        r#"Console.WriteLine(new[]{"x","y","z"}.MinBy(s=>s.Length));"#,
        ["x"]
    };

    min_by_on_single_element => {
        r#"Console.WriteLine(new[]{7}.MinBy(n=>n));"#,
        ["7"]
    };

    max_by_longest_word => {
        r#"Console.WriteLine(new[]{"a","bbb","cc"}.MaxBy(w=>w.Length));"#,
        ["bbb"]
    };

    max_by_longest_word_length => {
        r#"Console.WriteLine(new[]{"a","bbb","cc"}.MaxBy(w=>w.Length).Length);"#,
        ["3"]
    };

    max_by_highest_score => {
        r#"Console.WriteLine(new[]{(N:"a",S:1),(N:"b",S:5),(N:"c",S:3)}.MaxBy(t=>t.S).N);"#,
        ["b"]
    };

    max_by_last_char => {
        r#"Console.WriteLine(new[]{"cat","dog","cow"}.MaxBy(s=>s[s.Length-1]));"#,
        ["dog"]
    };

    max_by_on_single_element => {
        r#"Console.WriteLine(new[]{9}.MaxBy(n=>n));"#,
        ["9"]
    };

    element_at_zero_first => {
        r#"Console.WriteLine(new[]{10,20,30}.ElementAt(0));"#,
        ["10"]
    };

    element_at_one_middle => {
        r#"Console.WriteLine(new[]{10,20,30}.ElementAt(1));"#,
        ["20"]
    };

    element_at_last_index => {
        r#"Console.WriteLine(new[]{10,20,30}.ElementAt(2));"#,
        ["30"]
    };

    element_at_on_strings => {
        r#"Console.WriteLine(new[]{"a","b","c"}.ElementAt(1));"#,
        ["b"]
    };

    element_at_after_skip => {
        r#"Console.WriteLine(new[]{1,2,3,4,5}.Skip(2).ElementAt(1));"#,
        ["4"]
    };

    element_at_after_orderby => {
        r#"Console.WriteLine(new[]{3,1,2}.OrderBy(x=>x).ElementAt(1));"#,
        ["2"]
    };

    single_one_match => {
        r#"Console.WriteLine(new[]{42}.Single());"#,
        ["42"]
    };

    single_with_predicate_one_match => {
        r#"Console.WriteLine(new[]{1,2,3,4}.Single(x=>x==3));"#,
        ["3"]
    };

    single_with_predicate_zero_matches_caught => {
        r#"string tag="ok";
try{new[]{1,2,3}.Single(x=>x>10);}catch(System.InvalidOperationException){tag="none";}
Console.WriteLine(tag);"#,
        ["none"]
    };

    single_with_predicate_many_matches_caught => {
        r#"string tag="ok";
try{new[]{1,2,2}.Single(x=>x==2);}catch(System.InvalidOperationException){tag="many";}
Console.WriteLine(tag);"#,
        ["many"]
    };

    single_empty_throws_caught => {
        r#"string tag="ok";
try{System.Array.Empty<int>().Single();}catch(System.InvalidOperationException){tag="empty";}
Console.WriteLine(tag);"#,
        ["empty"]
    };

    single_two_elements_throws_caught => {
        r#"string tag="ok";
try{new[]{1,2}.Single();}catch(System.InvalidOperationException){tag="many";}
Console.WriteLine(tag);"#,
        ["many"]
    };

    single_or_default_one_element => {
        r#"Console.WriteLine(new[]{7}.SingleOrDefault());"#,
        ["7"]
    };

    single_or_default_empty_returns_default => {
        r#"Console.WriteLine(System.Array.Empty<int>().SingleOrDefault());"#,
        ["0"]
    };

    single_or_default_empty_with_seed => {
        r#"Console.WriteLine(System.Array.Empty<int>().SingleOrDefault(99));"#,
        ["99"]
    };

    single_or_default_many_returns_default => {
        r#"Console.WriteLine(new[]{1,2}.SingleOrDefault());"#,
        ["0"]
    };

    single_or_default_many_with_seed => {
        r#"Console.WriteLine(new[]{1,2}.SingleOrDefault(88));"#,
        ["88"]
    };

    single_or_default_predicate_one_match => {
        r#"Console.WriteLine(new[]{1,2,3}.SingleOrDefault(x=>x==2));"#,
        ["2"]
    };

    single_or_default_predicate_zero_returns_default => {
        r#"Console.WriteLine(new[]{1,2,3}.SingleOrDefault(x=>x>10));"#,
        ["0"]
    };

    single_or_default_predicate_zero_with_seed => {
        r#"Console.WriteLine(new[]{1,2,3}.SingleOrDefault(55,x=>x>10));"#,
        ["55"]
    };

    single_or_default_predicate_many_with_seed => {
        r#"Console.WriteLine(new[]{2,2,3}.SingleOrDefault(77,x=>x==2));"#,
        ["77"]
    };

    element_at_or_default_in_range => {
        r#"Console.WriteLine(new[]{5,6,7}.ElementAtOrDefault(1));"#,
        ["6"]
    };

    element_at_or_default_out_of_range => {
        r#"Console.WriteLine(new[]{5,6,7}.ElementAtOrDefault(10));"#,
        ["0"]
    };

    min_by_max_by_same_sequence_lengths => {
        r#"var words=new[]{"go","stop","run"};
Console.WriteLine(words.MinBy(w=>w.Length).Length);
Console.WriteLine(words.MaxBy(w=>w.Length).Length);"#,
        ["2", "4"]
    };

    aggregate_then_element_at => {
        r#"var running=new[]{1,2,3}.Aggregate(new int[]{0},(acc,x)=>new int[]{acc[0]+x});
Console.WriteLine(running[0]);"#,
        ["6"]
    };

    aggregate_seed_bool_all_true => {
        r#"var ok=new[]{true,true,true}.Aggregate(true,(acc,x)=>acc&&x);
Console.WriteLine(ok);"#,
        ["True"]
    };

    aggregate_seed_bool_one_false => {
        r#"var ok=new[]{true,false,true}.Aggregate(true,(acc,x)=>acc&&x);
Console.WriteLine(ok);"#,
        ["False"]
    };

    single_or_default_string_empty => {
        r#"Console.WriteLine(System.Array.Empty<string>().SingleOrDefault());"#,
        [""]
    };

    single_or_default_string_many_default => {
        r#"Console.WriteLine(new[]{"a","b"}.SingleOrDefault("z"));"#,
        ["z"]
    };

    element_at_large_index_throws_caught => {
        r#"string tag="ok";
try{new[]{1,2}.ElementAt(5);}catch(System.ArgumentOutOfRangeException){tag="range";}
Console.WriteLine(tag);"#,
        ["range"]
    };

    aggregate_no_seed_max_via_fold => {
        r#"Console.WriteLine(new[]{3,1,4}.Aggregate((a,b)=>a>b?a:b));"#,
        ["4"]
    };
}
