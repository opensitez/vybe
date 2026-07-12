//! StringBuilder gap coverage: Append/AppendLine/Insert/Remove/Replace/Clear/Length/Capacity overloads and edge cases.

csharp_cases! {
    stringbuilder_gap_append_char_single => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append('Z'); Console.WriteLine(sb.ToString());"#,
        ["Z"]
    };

    stringbuilder_gap_append_char_sequence => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append('a').Append('b').Append('c'); Console.WriteLine(sb.ToString());"#,
        ["abc"]
    };

    stringbuilder_gap_append_int_positive => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(42); Console.WriteLine(sb.ToString());"#,
        ["42"]
    };

    stringbuilder_gap_append_int_negative => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(-7); Console.WriteLine(sb.ToString());"#,
        ["-7"]
    };

    stringbuilder_gap_append_bool_true => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(true); Console.WriteLine(sb.ToString());"#,
        ["True"]
    };

    stringbuilder_gap_append_bool_false => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(false); Console.WriteLine(sb.ToString());"#,
        ["False"]
    };

    stringbuilder_gap_append_decimal_literal => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(3.5m); Console.WriteLine(sb.ToString());"#,
        ["3.5"]
    };

    stringbuilder_gap_append_double_value => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(2.5); Console.WriteLine(sb.ToString());"#,
        ["2.5"]
    };

    stringbuilder_gap_append_long_value => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append(10000000000L); Console.WriteLine(sb.ToString());"#,
        ["10000000000"]
    };

    stringbuilder_gap_append_empty_string => {
        r#"var sb=new System.Text.StringBuilder("x"); sb.Append(""); Console.WriteLine(sb.ToString());"#,
        ["x"]
    };

    stringbuilder_gap_appendline_no_argument => {
        r#"var sb=new System.Text.StringBuilder("a"); sb.AppendLine(); Console.WriteLine(sb.Length);"#,
        ["3"]
    };

    stringbuilder_gap_appendline_char => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendLine('x'); Console.WriteLine(sb.ToString().Trim());"#,
        ["x"]
    };

    stringbuilder_gap_appendline_int => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendLine(9); Console.WriteLine(sb.ToString().Trim());"#,
        ["9"]
    };

    stringbuilder_gap_appendline_chained_three => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendLine("a").AppendLine("b").AppendLine("c"); Console.WriteLine(sb.ToString().Replace("\r\n","\n").Trim().Split('\n').Length);"#,
        ["3"]
    };

    stringbuilder_gap_insert_at_zero => {
        r#"var sb=new System.Text.StringBuilder("bc"); sb.Insert(0,"a"); Console.WriteLine(sb.ToString());"#,
        ["abc"]
    };

    stringbuilder_gap_insert_at_end => {
        r#"var sb=new System.Text.StringBuilder("ab"); sb.Insert(2,"c"); Console.WriteLine(sb.ToString());"#,
        ["abc"]
    };

    stringbuilder_gap_insert_char_at_middle => {
        r#"var sb=new System.Text.StringBuilder("ac"); sb.Insert(1,'b'); Console.WriteLine(sb.ToString());"#,
        ["abc"]
    };

    stringbuilder_gap_insert_int_at_start => {
        r#"var sb=new System.Text.StringBuilder("end"); sb.Insert(0,1); Console.WriteLine(sb.ToString());"#,
        ["1end"]
    };

    stringbuilder_gap_insert_empty_string_noop => {
        r#"var sb=new System.Text.StringBuilder("same"); sb.Insert(2,""); Console.WriteLine(sb.ToString());"#,
        ["same"]
    };

    stringbuilder_gap_insert_multiple_times => {
        r#"var sb=new System.Text.StringBuilder("a"); sb.Insert(1,"b").Insert(2,"c"); Console.WriteLine(sb.ToString());"#,
        ["abc"]
    };

    stringbuilder_gap_remove_first_character => {
        r#"var sb=new System.Text.StringBuilder("abc"); sb.Remove(0,1); Console.WriteLine(sb.ToString());"#,
        ["bc"]
    };

    stringbuilder_gap_remove_last_character => {
        r#"var sb=new System.Text.StringBuilder("abc"); sb.Remove(2,1); Console.WriteLine(sb.ToString());"#,
        ["ab"]
    };

    stringbuilder_gap_remove_middle_range => {
        r#"var sb=new System.Text.StringBuilder("abcde"); sb.Remove(1,3); Console.WriteLine(sb.ToString());"#,
        ["ae"]
    };

    stringbuilder_gap_remove_all_content => {
        r#"var sb=new System.Text.StringBuilder("hello"); sb.Remove(0,5); Console.WriteLine(sb.Length);"#,
        ["0"]
    };

    stringbuilder_gap_remove_zero_count => {
        r#"var sb=new System.Text.StringBuilder("keep"); sb.Remove(2,0); Console.WriteLine(sb.ToString());"#,
        ["keep"]
    };

    stringbuilder_gap_replace_single_occurrence => {
        r#"var sb=new System.Text.StringBuilder("cat"); sb.Replace("a","o"); Console.WriteLine(sb.ToString());"#,
        ["cot"]
    };

    stringbuilder_gap_replace_all_occurrences => {
        r#"var sb=new System.Text.StringBuilder("banana"); sb.Replace("a","o"); Console.WriteLine(sb.ToString());"#,
        ["bonono"]
    };

    stringbuilder_gap_replace_no_match_unchanged => {
        r#"var sb=new System.Text.StringBuilder("xyz"); sb.Replace("q","w"); Console.WriteLine(sb.ToString());"#,
        ["xyz"]
    };

    stringbuilder_gap_replace_shorter_with_longer => {
        r#"var sb=new System.Text.StringBuilder("a-b"); sb.Replace("-","->"); Console.WriteLine(sb.ToString());"#,
        ["a->b"]
    };

    stringbuilder_gap_replace_longer_with_shorter => {
        r#"var sb=new System.Text.StringBuilder("a->b"); sb.Replace("->","-"); Console.WriteLine(sb.ToString());"#,
        ["a-b"]
    };

    stringbuilder_gap_replace_char_pair => {
        r#"var sb=new System.Text.StringBuilder("x1x2"); sb.Replace('x','y'); Console.WriteLine(sb.ToString());"#,
        ["y1y2"]
    };

    stringbuilder_gap_replace_after_insert => {
        r#"var sb=new System.Text.StringBuilder("ab"); sb.Insert(1,"X"); sb.Replace("X","-"); Console.WriteLine(sb.ToString());"#,
        ["a-b"]
    };

    stringbuilder_gap_clear_on_empty_builder => {
        r#"var sb=new System.Text.StringBuilder(); sb.Clear(); Console.WriteLine(sb.Length);"#,
        ["0"]
    };

    stringbuilder_gap_clear_then_append => {
        r#"var sb=new System.Text.StringBuilder("old"); sb.Clear(); sb.Append("new"); Console.WriteLine(sb.ToString());"#,
        ["new"]
    };

    stringbuilder_gap_clear_twice => {
        r#"var sb=new System.Text.StringBuilder("data"); sb.Clear(); sb.Clear(); Console.WriteLine(sb.Length);"#,
        ["0"]
    };

    stringbuilder_gap_length_after_append => {
        r#"var sb=new System.Text.StringBuilder(); sb.Append("four"); Console.WriteLine(sb.Length);"#,
        ["4"]
    };

    stringbuilder_gap_length_after_appendline => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendLine("x"); Console.WriteLine(sb.Length>=2);"#,
        ["True"]
    };

    stringbuilder_gap_length_after_remove => {
        r#"var sb=new System.Text.StringBuilder("abcdef"); sb.Remove(2,2); Console.WriteLine(sb.Length);"#,
        ["4"]
    };

    stringbuilder_gap_length_after_replace_growth => {
        r#"var sb=new System.Text.StringBuilder("a"); sb.Replace("a","long"); Console.WriteLine(sb.Length);"#,
        ["4"]
    };

    stringbuilder_gap_capacity_default_constructor => {
        r#"var sb=new System.Text.StringBuilder(); Console.WriteLine(sb.Capacity>=16);"#,
        ["True"]
    };

    stringbuilder_gap_capacity_with_initial_string => {
        r#"var sb=new System.Text.StringBuilder("hello"); Console.WriteLine(sb.Capacity>=5);"#,
        ["True"]
    };

    stringbuilder_gap_capacity_explicit_small => {
        r#"var sb=new System.Text.StringBuilder(8); Console.WriteLine(sb.Capacity);"#,
        ["8"]
    };

    stringbuilder_gap_capacity_grows_after_many_appends => {
        r#"var sb=new System.Text.StringBuilder(4); for(int i=0;i<50;i++) sb.Append('q'); Console.WriteLine(sb.Capacity>=50);"#,
        ["True"]
    };

    stringbuilder_gap_capacity_set_larger => {
        r#"var sb=new System.Text.StringBuilder("hi"); sb.Capacity=64; Console.WriteLine(sb.Capacity>=64);"#,
        ["True"]
    };

    stringbuilder_gap_capacity_unchanged_when_content_fits => {
        r#"var sb=new System.Text.StringBuilder(32); sb.Append("tiny"); Console.WriteLine(sb.Capacity);"#,
        ["32"]
    };

    stringbuilder_gap_appendformat_two_placeholders => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendFormat("{0}-{1}",1,2); Console.WriteLine(sb.ToString());"#,
        ["1-2"]
    };

    stringbuilder_gap_appendformat_with_literal_braces => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendFormat("{{0}}={0}",5); Console.WriteLine(sb.ToString());"#,
        ["{0}=5"]
    };

    stringbuilder_gap_indexer_after_mutations => {
        r#"var sb=new System.Text.StringBuilder("abc"); sb[1]='B'; sb.Append("d"); Console.WriteLine(sb[0]); Console.WriteLine(sb.ToString());"#,
        ["a", "aBcd"]
    };

    stringbuilder_gap_tostring_empty_builder => {
        r#"var sb=new System.Text.StringBuilder(); Console.WriteLine(sb.ToString()=="");"#,
        ["True"]
    };

    stringbuilder_gap_tostring_after_clear => {
        r#"var sb=new System.Text.StringBuilder("gone"); sb.Clear(); Console.WriteLine(sb.ToString());"#,
        [""]
    };

    stringbuilder_gap_mixed_append_insert_remove => {
        r#"var sb=new System.Text.StringBuilder("start"); sb.Append("_end"); sb.Insert(5,"-mid-"); sb.Remove(0,6); Console.WriteLine(sb.ToString());"#,
        ["mid-_end"]
    };

    stringbuilder_gap_appendline_then_replace_newline => {
        r#"var sb=new System.Text.StringBuilder(); sb.AppendLine("row"); Console.WriteLine(sb.ToString().Contains("\n")||sb.ToString().Contains("\r"));"#,
        ["True"]
    };

    stringbuilder_gap_constructor_with_capacity_then_seed => {
        r#"var sb=new System.Text.StringBuilder(16); sb.Append("seed"); sb.Append("+"); Console.WriteLine(sb.ToString());"#,
        ["seed+"]
    };
}
