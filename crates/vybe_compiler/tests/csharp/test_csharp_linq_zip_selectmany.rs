//! LINQ `Zip` and `SelectMany`: pairing sequences and flattening nested collections.
//! GAP: Zip/SelectMany depth beyond existing `test_csharp_linq_*` coverage.

csharp_cases! {
    zip_multiply_pairs_returns_product_sequence_count => {
        r#"var z=new[]{1,2,3}.Zip(new[]{10,20,30},(a,b)=>a*b);
Console.WriteLine(z.Count());"#,
        ["3"]
    };

    zip_multiply_pairs_first_product => {
        r#"var z=new[]{1,2,3}.Zip(new[]{10,20,30},(a,b)=>a*b);
Console.WriteLine(z.First());"#,
        ["10"]
    };

    zip_multiply_pairs_last_product => {
        r#"var z=new[]{1,2,3}.Zip(new[]{10,20,30},(a,b)=>a*b);
Console.WriteLine(z.Last());"#,
        ["90"]
    };

    zip_multiply_pairs_sum => {
        r#"var z=new[]{1,2,3}.Zip(new[]{10,20,30},(a,b)=>a*b);
Console.WriteLine(z.Sum());"#,
        ["140"]
    };

    zip_unequal_lengths_stops_at_shorter_count => {
        r#"var z=new[]{1,2,3,4,5}.Zip(new[]{10,20},(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["2"]
    };

    zip_unequal_lengths_second_shorter_count => {
        r#"var z=new[]{1,2}.Zip(new[]{10,20,30},(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["2"]
    };

    zip_empty_first_sequence_yields_zero_pairs => {
        r#"var z=System.Array.Empty<int>().Zip(new[]{1,2,3},(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["0"]
    };

    zip_empty_second_sequence_yields_zero_pairs => {
        r#"var z=new[]{1,2,3}.Zip(System.Array.Empty<int>(),(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["0"]
    };

    zip_both_empty_yields_zero_pairs => {
        r#"var z=System.Array.Empty<int>().Zip(System.Array.Empty<int>(),(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["0"]
    };

    zip_string_pairs_concatenated_count => {
        r#"var z=new[]{"a","b"}.Zip(new[]{"1","2"},(x,y)=>x+y);
Console.WriteLine(z.Count());"#,
        ["2"]
    };

    zip_string_pairs_first_concatenation => {
        r#"var z=new[]{"a","b"}.Zip(new[]{"1","2"},(x,y)=>x+y);
Console.WriteLine(z.First());"#,
        ["a1"]
    };

    zip_string_pairs_second_concatenation => {
        r#"var z=new[]{"a","b"}.Zip(new[]{"1","2"},(x,y)=>x+y);
Console.WriteLine(z.Last());"#,
        ["b2"]
    };

    zip_single_element_pair_sum => {
        r#"var z=new[]{7}.Zip(new[]{5},(a,b)=>a+b);
Console.WriteLine(z.Single());"#,
        ["12"]
    };

    zip_result_selector_returns_tuple_like_sum => {
        r#"var z=new[]{1,2}.Zip(new[]{3,4},(a,b)=>a*10+b);
Console.WriteLine(z.Sum());"#,
        ["47"]
    };

    zip_bool_sequences_and_count => {
        r#"var z=new[]{true,false,true}.Zip(new[]{false,true,false},(a,b)=>a&&b);
Console.WriteLine(z.Count(x=>x));"#,
        ["0"]
    };

    zip_char_sequences_count => {
        r#"var z=new[]{'a','b'}.Zip(new[]{'x','y'},(a,b)=>(char)(a+b-'a'+'x'));
Console.WriteLine(z.Count());"#,
        ["2"]
    };

    zip_after_skip_on_first_side_count => {
        r#"var z=new[]{1,2,3,4}.Skip(1).Zip(new[]{10,20,30},(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["3"]
    };

    zip_after_take_on_second_side_count => {
        r#"var z=new[]{1,2,3}.Zip(new[]{10,20,30,40}.Take(2),(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["2"]
    };

    zip_enumerate_products_in_foreach => {
        r#"var z=new[]{2,3}.Zip(new[]{4,5},(a,b)=>a*b);
foreach(var n in z) Console.WriteLine(n);"#,
        ["8", "15"]
    };

    zip_three_way_manual_via_select_index => {
        r#"var a=new[]{1,2,3}; var b=new[]{4,5,6};
var z=a.Zip(b,(x,y)=>x+y);
Console.WriteLine(z.ElementAt(1));"#,
        ["7"]
    };

    zip_preserves_left_order_first => {
        r#"var z=new[]{3,1,2}.Zip(new[]{1,1,1},(a,b)=>a);
Console.WriteLine(z.First());"#,
        ["3"]
    };

    zip_negative_numbers_product_count => {
        r#"var z=new[]{-1,2}.Zip(new[]{3,-4},(a,b)=>a*b);
Console.WriteLine(z.Count());"#,
        ["2"]
    };

    zip_doubles_sum => {
        r#"var z=new[]{1.5,2.5}.Zip(new[]{2.0,2.0},(a,b)=>a+b);
Console.WriteLine(z.Sum());"#,
        ["8"]
    };

    select_many_flatten_jagged_arrays_count => {
        r#"var flat=new[]{new[]{1,2},new[]{3,4,5}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["5"]
    };

    select_many_flatten_jagged_arrays_sum => {
        r#"var flat=new[]{new[]{1,2},new[]{3,4,5}}.SelectMany(x=>x);
Console.WriteLine(flat.Sum());"#,
        ["15"]
    };

    select_many_three_nested_levels_count => {
        r#"var data=new[]{new[]{new[]{1,2}},new[]{new[]{3}}};
var flat=data.SelectMany(a=>a).SelectMany(b=>b);
Console.WriteLine(flat.Count());"#,
        ["3"]
    };

    select_many_three_nested_levels_sum => {
        r#"var data=new[]{new[]{new[]{1,2}},new[]{new[]{3}}};
var flat=data.SelectMany(a=>a).SelectMany(b=>b);
Console.WriteLine(flat.Sum());"#,
        ["6"]
    };

    select_many_empty_inner_sequences_count => {
        r#"var flat=new[]{new int[]{},new[]{1,2},new int[]{}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["2"]
    };

    select_many_all_empty_inner_sequences_count => {
        r#"var flat=new[]{new int[]{},new int[]{}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["0"]
    };

    select_many_single_element_inner_count => {
        r#"var flat=new[]{new[]{1},new[]{2},new[]{3}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["3"]
    };

    select_many_strings_to_chars_count => {
        r#"var chars=new[]{"ab","c"}.SelectMany(s=>s);
Console.WriteLine(chars.Count());"#,
        ["3"]
    };

    select_many_strings_to_chars_first => {
        r#"var chars=new[]{"ab","c"}.SelectMany(s=>s);
Console.WriteLine(chars.First());"#,
        ["a"]
    };

    select_many_strings_to_chars_last => {
        r#"var chars=new[]{"ab","c"}.SelectMany(s=>s);
Console.WriteLine(chars.Last());"#,
        ["c"]
    };

    select_many_with_result_selector_count => {
        r#"var flat=new[]{new[]{1,2},new[]{3}}.SelectMany(x=>x,y=>y*10);
Console.WriteLine(flat.Count());"#,
        ["3"]
    };

    select_many_with_result_selector_sum => {
        r#"var flat=new[]{new[]{1,2},new[]{3}}.SelectMany(x=>x,y=>y*10);
Console.WriteLine(flat.Sum());"#,
        ["60"]
    };

    select_many_with_index_count => {
        r#"var flat=new[]{new[]{10},new[]{20,30}}.SelectMany((x,i)=>x.Select(v=>v+i));
Console.WriteLine(flat.Count());"#,
        ["3"]
    };

    select_many_with_index_first_value => {
        r#"var flat=new[]{new[]{10},new[]{20,30}}.SelectMany((x,i)=>x.Select(v=>v+i));
Console.WriteLine(flat.First());"#,
        ["10"]
    };

    select_many_with_index_last_value => {
        r#"var flat=new[]{new[]{10},new[]{20,30}}.SelectMany((x,i)=>x.Select(v=>v+i));
Console.WriteLine(flat.Last());"#,
        ["31"]
    };

    select_many_from_list_of_lists_count => {
        r#"var lists=new System.Collections.Generic.List<int[]>{
    new[]{1,2},new[]{3}}; 
Console.WriteLine(lists.SelectMany(x=>x).Count());"#,
        ["3"]
    };

    select_many_preserves_order_foreach => {
        r#"var flat=new[]{new[]{1,2},new[]{3,4}}.SelectMany(x=>x);
foreach(var n in flat) Console.WriteLine(n);"#,
        ["1", "2", "3", "4"]
    };

    select_many_after_where_on_outer_count => {
        r#"var flat=new[]{new[]{1,2},new[]{3,4}}.Where(a=>a.Length>1).SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["2"]
    };

    select_many_identity_on_scalars_via_array_return => {
        r#"var flat=new[]{1,2,3}.SelectMany(n=>new[]{n,n});
Console.WriteLine(flat.Count());"#,
        ["6"]
    };

    select_many_expand_each_to_range_count => {
        r#"var flat=new[]{1,2,3}.SelectMany(n=>Enumerable.Range(1,n));
Console.WriteLine(flat.Count());"#,
        ["6"]
    };

    select_many_expand_each_to_range_sum => {
        r#"var flat=new[]{1,2,3}.SelectMany(n=>Enumerable.Range(1,n));
Console.WriteLine(flat.Sum());"#,
        ["10"]
    };

    select_many_empty_source_count => {
        r#"var flat=System.Array.Empty<int[]>().SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["0"]
    };

    select_many_mixed_empty_and_nonempty_count => {
        r#"var flat=new[]{new int[]{},new[]{5,6},new int[]{}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["2"]
    };

    select_many_large_inner_sequence_count => {
        r#"var flat=new[]{new[]{1,2,3,4,5}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["5"]
    };

    select_many_doubles_count => {
        r#"var flat=new[]{new[]{1.5,2.5},new[]{3.0}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["3"]
    };

    select_many_zip_then_select_many_count => {
        r#"var pairs=new[]{1,2}.Zip(new[]{3,4},(a,b)=>new[]{a,b});
Console.WriteLine(pairs.SelectMany(x=>x).Count());"#,
        ["4"]
    };

    select_many_zip_then_select_many_sum => {
        r#"var pairs=new[]{1,2}.Zip(new[]{3,4},(a,b)=>new[]{a,b});
Console.WriteLine(pairs.SelectMany(x=>x).Sum());"#,
        ["10"]
    };

    zip_then_select_many_char_count => {
        r#"var words=new[]{"hi","go"};
var letters=words.Zip(new[]{1,2},(w,n)=>w).SelectMany(w=>w);
Console.WriteLine(letters.Count());"#,
        ["4"]
    };

    select_many_nested_string_arrays_joined_count => {
        r#"var flat=new[]{new[]{"a","b"},new[]{"c"}}.SelectMany(x=>x);
Console.WriteLine(flat.Count());"#,
        ["3"]
    };

    select_many_nested_string_arrays_joined_foreach => {
        r#"var flat=new[]{new[]{"a","b"},new[]{"c"}}.SelectMany(x=>x);
foreach(var s in flat) Console.WriteLine(s);"#,
        ["a", "b", "c"]
    };

    zip_longer_first_with_take_count => {
        r#"var z=new[]{1,2,3,4}.Take(3).Zip(new[]{10,20,30,40},(a,b)=>a+b);
Console.WriteLine(z.Count());"#,
        ["3"]
    };

    select_many_repeat_each_element_count => {
        r#"var flat=new[]{1,2}.SelectMany(n=>new[]{n,n,n});
Console.WriteLine(flat.Count());"#,
        ["6"]
    };

    select_many_repeat_each_element_sum => {
        r#"var flat=new[]{1,2}.SelectMany(n=>new[]{n,n,n});
Console.WriteLine(flat.Sum());"#,
        ["9"]
    };
}
