//! List patterns: `[a,b]`, discard `_`, `var` positional slots, slices, and switch arms.

csharp_cases! {
    is_list_empty_pattern_matches_zero_length_array => {
        r#"int[] data = new int[]{}; Console.WriteLine(data is []);"#,
        ["True"]
    };

    is_list_empty_pattern_rejects_non_empty_array => {
        r#"int[] data = new[]{1}; Console.WriteLine(data is []);"#,
        ["False"]
    };

    is_list_single_var_pattern_captures_element => {
        r#"int[] data = new[]{42}; if (data is [var n]) Console.WriteLine(n);"#,
        ["42"]
    };

    is_list_single_discard_pattern_accepts_one_element => {
        r#"int[] data = new[]{7}; Console.WriteLine(data is [_]);"#,
        ["True"]
    };

    is_list_single_var_pattern_rejects_empty_array => {
        r#"int[] data = new int[]{}; Console.WriteLine(data is [var n]);"#,
        ["False"]
    };

    is_list_pair_constant_pattern_matches_exact_sequence => {
        r#"int[] data = new[]{1,2}; Console.WriteLine(data is [1,2]);"#,
        ["True"]
    };

    is_list_pair_constant_pattern_rejects_wrong_values => {
        r#"int[] data = new[]{1,3}; Console.WriteLine(data is [1,2]);"#,
        ["False"]
    };

    is_list_pair_var_pattern_binds_both_positions => {
        r#"int[] data = new[]{3,4}; if (data is [var a,var b]) Console.WriteLine(a+b);"#,
        ["7"]
    };

    is_list_pair_discard_pattern_checks_length_only => {
        r#"int[] data = new[]{9,1}; Console.WriteLine(data is [_,_]);"#,
        ["True"]
    };

    is_list_pair_discard_pattern_rejects_single_element => {
        r#"int[] data = new[]{9}; Console.WriteLine(data is [_,_]);"#,
        ["False"]
    };

    is_list_triple_var_pattern_destructures_three_slots => {
        r#"int[] data = new[]{1,2,3}; if (data is [var a,var b,var c]) Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    is_list_triple_constant_pattern_matches_literal_head => {
        r#"int[] data = new[]{5,6,7}; Console.WriteLine(data is [5,6,7]);"#,
        ["True"]
    };

    is_list_first_constant_second_var_pattern => {
        r#"int[] data = new[]{0,15}; if (data is [0,var tail]) Console.WriteLine(tail);"#,
        ["15"]
    };

    is_list_first_var_second_constant_pattern => {
        r#"int[] data = new[]{8,2}; if (data is [var head,2]) Console.WriteLine(head);"#,
        ["8"]
    };

    is_list_mixed_discard_and_var_pattern => {
        r#"int[] data = new[]{11,22,33}; if (data is [var a,_,var c]) Console.WriteLine(a+c);"#,
        ["44"]
    };

    is_list_slice_rest_captures_tail_length => {
        r#"int[] data = new[]{1,2,3,4}; if (data is [var head,..,var last]) Console.WriteLine(last-head);"#,
        ["3"]
    };

    is_list_slice_open_start_matches_suffix_constant => {
        r#"int[] data = new[]{9,8,7}; Console.WriteLine(data is [..,7]);"#,
        ["True"]
    };

    is_list_slice_open_end_matches_prefix_constant => {
        r#"int[] data = new[]{9,8,7}; Console.WriteLine(data is [9,..]);"#,
        ["True"]
    };

    is_list_slice_bookended_constants_match_middle_gap => {
        r#"int[] data = new[]{1,2,3,4,5}; Console.WriteLine(data is [1,..,5]);"#,
        ["True"]
    };

    is_list_slice_bookended_constants_reject_wrong_ends => {
        r#"int[] data = new[]{1,2,3,4,6}; Console.WriteLine(data is [1,..,5]);"#,
        ["False"]
    };

    is_list_slice_head_var_rest_discard => {
        r#"int[] data = new[]{4,5,6}; if (data is [var first,..]) Console.WriteLine(first);"#,
        ["4"]
    };

    is_list_slice_discard_head_tail_var => {
        r#"int[] data = new[]{4,5,6}; if (data is [..,var last]) Console.WriteLine(last);"#,
        ["6"]
    };

    is_list_slice_on_single_element_has_empty_rest => {
        r#"int[] data = new[]{99}; if (data is [var a,..var rest]) Console.WriteLine(rest.Length);"#,
        ["0"]
    };

    is_list_slice_on_pair_splits_head_and_tail => {
        r#"int[] data = new[]{5,6}; if (data is [var a,..var rest]) Console.WriteLine(a+rest[0]);"#,
        ["11"]
    };

    switch_expression_list_empty_arm_labels_empty => {
        r#"string Label(int[] a)=>a switch{[]=>"empty",_=>"other"}; Console.WriteLine(Label(new int[]{}));"#,
        ["empty"]
    };

    switch_expression_list_single_discard_arm_labels_single => {
        r#"string Label(int[] a)=>a switch{[_]=>"one",_=>"other"}; Console.WriteLine(Label(new[]{9}));"#,
        ["one"]
    };

    switch_expression_list_pair_discard_arm_labels_pair => {
        r#"string Label(int[] a)=>a switch{[_,_]=>"pair",_=>"other"}; Console.WriteLine(Label(new[]{1,2}));"#,
        ["pair"]
    };

    switch_expression_list_triple_discard_arm_labels_triple => {
        r#"string Label(int[] a)=>a switch{[_,_,_]=>"triple",_=>"other"}; Console.WriteLine(Label(new[]{1,2,3}));"#,
        ["triple"]
    };

    switch_expression_list_var_pair_arm_returns_sum => {
        r#"int SumPair(int[] a)=>a switch{[var x,var y]=>x+y,_=>0}; Console.WriteLine(SumPair(new[]{10,20}));"#,
        ["30"]
    };

    switch_expression_list_constant_pair_arm_matches_literals => {
        r#"string Code(int[] a)=>a switch{[1,2]=>"twelve",_=>"other"}; Console.WriteLine(Code(new[]{1,2}));"#,
        ["twelve"]
    };

    switch_expression_list_many_arm_after_fixed_lengths => {
        r#"string Size(int[] a)=>a switch{[]=>"0",[_]=>"1",[_,_]=>"2",_=>"many"}; Console.WriteLine(Size(new[]{1,2,3}));"#,
        ["many"]
    };

    switch_expression_list_slice_arm_checks_bookends => {
        r#"string Edge(int[] a)=>a switch{[1,..,9]=>"book",_=>"plain"}; Console.WriteLine(Edge(new[]{1,5,9}));"#,
        ["book"]
    };

    switch_statement_list_pattern_case_matches_triple => {
        r#"int[] data=new[]{2,4,6}; string tag=""; switch(data){case[2,4,6]:tag="hit";break;default:tag="miss";break;} Console.WriteLine(tag);"#,
        ["hit"]
    };

    switch_statement_list_pattern_case_with_var_capture => {
        r#"int[] data=new[]{3,9}; string tag=""; switch(data){case[var a,var b]:tag=(a+b).ToString();break;default:tag="0";break;} Console.WriteLine(tag);"#,
        ["12"]
    };

    is_list_string_array_exact_sequence => {
        r#"string[] words=new[]{"a","b"}; Console.WriteLine(words is ["a","b"]);"#,
        ["True"]
    };

    is_list_string_var_pattern_captures_element => {
        r#"string[] words=new[]{"hi"}; if(words is [var w]) Console.WriteLine(w);"#,
        ["hi"]
    };

    is_list_bool_pair_constants => {
        r#"bool[] flags=new[]{true,false}; Console.WriteLine(flags is [true,false]);"#,
        ["True"]
    };

    is_list_double_triple_sum_via_vars => {
        r#"double[] vals=new[]{1.5,2.0,2.5}; if(vals is [var a,var b,var c]) Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    is_list_byte_array_length_two_discard => {
        r#"byte[] buf=new byte[]{10,20}; Console.WriteLine(buf is [_,_]);"#,
        ["True"]
    };

    is_list_long_array_single_capture => {
        r#"long[] ids=new long[]{1000L}; if(ids is [var id]) Console.WriteLine(id);"#,
        ["1000"]
    };

    is_list_not_pattern_inverts_match => {
        r#"int[] data=new[]{1,2}; Console.WriteLine(data is not [1,2]);"#,
        ["False"]
    };

    is_list_not_pattern_accepts_different_shape => {
        r#"int[] data=new[]{1}; Console.WriteLine(data is not [1,2]);"#,
        ["True"]
    };

    // `when` after `is` is not valid C# (guards exist only on `case`/`catch`);
    // the equivalent is `&&`, where the pattern variables are in scope.
    is_list_pattern_with_guard_on_captured_var => {
        r#"int[] data=new[]{4,8}; if(data is [var a,var b] && a<b) Console.WriteLine("ordered");"#,
        ["ordered"]
    };

    is_list_pattern_guard_rejects_wrong_order => {
        r#"int[] data=new[]{8,4}; if(data is [var a,var b] && a<b) Console.WriteLine("ordered"); else Console.WriteLine("not");"#,
        ["not"]
    };

    switch_expression_list_when_guard_on_vars => {
        r#"string Rank(int[] a)=>a switch{[var x,var y] when x>y=>"desc",[var x,var y]=>"asc",_=>"other"}; Console.WriteLine(Rank(new[]{5,2}));"#,
        ["desc"]
    };

    switch_expression_list_when_guard_falls_to_second_arm => {
        r#"string Rank(int[] a)=>a switch{[var x,var y] when x>y=>"desc",[var x,var y]=>"asc",_=>"other"}; Console.WriteLine(Rank(new[]{2,5}));"#,
        ["asc"]
    };

    is_list_four_element_all_discard => {
        r#"int[] data=new[]{1,2,3,4}; Console.WriteLine(data is [_,_,_,_]);"#,
        ["True"]
    };

    is_list_four_element_rejects_three_length => {
        r#"int[] data=new[]{1,2,3}; Console.WriteLine(data is [_,_,_,_]);"#,
        ["False"]
    };

    is_list_zero_in_first_slot_constant => {
        r#"int[] data=new[]{0,42}; if(data is [0,var v]) Console.WriteLine(v);"#,
        ["42"]
    };

    is_list_negative_constants_in_pattern => {
        r#"int[] data=new[]{-1,-2}; Console.WriteLine(data is [-1,-2]);"#,
        ["True"]
    };

    switch_expression_list_returns_string_from_var_slots => {
        r#"string PairLabel(int[] a)=>a switch{[var x,var y]=>$"{x}-{y}",_=>"?"}; Console.WriteLine(PairLabel(new[]{7,8}));"#,
        ["7-8"]
    };

    is_list_collection_expression_literal_pair => {
        r#"int[] data=[10,20]; Console.WriteLine(data is [10,20]);"#,
        ["True"]
    };

    switch_expression_list_default_after_length_arms => {
        r#"string Bucket(int[] a)=>a switch{[]=>"e",[_]=>"s",_=>"m"}; Console.WriteLine(Bucket(new[]{1,2}));"#,
        ["m"]
    };

    is_list_char_array_single_element_capture => {
        r#"char[] chars=new[]{'x'}; if(chars is [var ch]) Console.WriteLine(ch);"#,
        ["x"]
    };

    is_list_int_array_exact_three_constants => {
        r#"int[] data=new[]{2,4,6}; Console.WriteLine(data is [2,4,6]);"#,
        ["True"]
    };
}
