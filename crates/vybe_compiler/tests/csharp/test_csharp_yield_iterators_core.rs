//! `IEnumerable<T>` iterator methods — `yield return` / `yield break`, nesting, and `finally` cleanup prints.

csharp_cases! {
    yield_return_three_values_sum_in_foreach => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){yield return 1;yield return 2;yield return 3;}
int s=0; foreach(var n in Gen()) s+=n; Console.WriteLine(s);"#,
        ["6"]
    };

    yield_return_single_value_count_one => {
        r#"System.Collections.Generic.IEnumerable<int> One(){yield return 42;}
int c=0; foreach(var _ in One()) c++; Console.WriteLine(c);"#,
        ["1"]
    };

    yield_break_before_second_yield_yields_one => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){yield return 1;yield break;yield return 2;}
int c=0; foreach(var _ in Gen()) c++; Console.WriteLine(c);"#,
        ["1"]
    };

    yield_break_in_loop_stops_at_limit => {
        r#"System.Collections.Generic.IEnumerable<int> Take(int max){for(int i=0;i<10;i++){if(i>=max)yield break;yield return i;}}
Console.WriteLine(string.Join(",",Take(3)));"#,
        ["0,1,2"]
    };

    yield_return_in_while_produces_infinite_prefix => {
        r#"System.Collections.Generic.IEnumerable<int> Up(){int n=0;while(n<4){yield return n;n++;}}
Console.WriteLine(string.Join(",",Up()));"#,
        ["0,1,2,3"]
    };

    yield_return_in_if_branch_only_when_true => {
        r#"System.Collections.Generic.IEnumerable<int> Pick(bool ok){if(ok){yield return 7;}yield return 0;}
Console.WriteLine(string.Join(",",Pick(true)));"#,
        ["7,0"]
    };

    yield_return_in_if_branch_skips_when_false => {
        r#"System.Collections.Generic.IEnumerable<int> Pick(bool ok){if(ok){yield return 7;}yield return 0;}
Console.WriteLine(string.Join(",",Pick(false)));"#,
        ["0"]
    };

    yield_return_in_switch_case_arms => {
        r#"System.Collections.Generic.IEnumerable<string> Label(int n){switch(n){case 1:yield return "one";break;case 2:yield return "two";break;default:yield return "many";break;}}
Console.WriteLine(string.Join("|",Label(2)));"#,
        ["two"]
    };

    yield_return_from_instance_method_on_class => {
        r#"class Counter{public System.Collections.Generic.IEnumerable<int> Range(int n){for(int i=0;i<n;i++)yield return i;}}
Console.WriteLine(new Counter().Range(4).Sum());"#,
        ["6"]
    };

    yield_return_from_static_method => {
        r#"class Seq{public static System.Collections.Generic.IEnumerable<int> Twice(int n){yield return n;yield return n*2;}}
Console.WriteLine(string.Join(",",Seq.Twice(5)));"#,
        ["5,10"]
    };

    yield_return_generic_type_parameter => {
        r#"System.Collections.Generic.IEnumerable<T> Echo<T>(T v){yield return v;yield return v;}
Console.WriteLine(string.Join(",",Echo("x")));"#,
        ["x,x"]
    };

    nested_iterator_yields_from_inner_foreach => {
        r#"System.Collections.Generic.IEnumerable<int> Inner(){yield return 1;yield return 2;}
System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in Inner())yield return x;foreach(var x in Inner())yield return x;}
Console.WriteLine(string.Join(",",Outer()));"#,
        ["1,2,1,2"]
    };

    nested_iterator_select_many_style_flatten => {
        r#"System.Collections.Generic.IEnumerable<int> Pair(int n){yield return n;yield return n+1;}
Console.WriteLine(string.Join(",",new[]{1,2}.SelectMany(Pair)));"#,
        ["1,2,2,3"]
    };

    nested_three_level_iterator_chain => {
        r#"System.Collections.Generic.IEnumerable<int> A(){yield return 1;}
System.Collections.Generic.IEnumerable<int> B(){foreach(var x in A())yield return x+10;}
System.Collections.Generic.IEnumerable<int> C(){foreach(var x in B())yield return x+100;}
Console.WriteLine(string.Join(",",C()));"#,
        ["111"]
    };

    iterator_finally_prints_after_full_consumption => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield return 2;}finally{Console.WriteLine("fin");}}
foreach(var _ in Gen()){} "#,
        ["fin"]
    };

    iterator_finally_prints_after_yield_break => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield break;yield return 9;}finally{Console.WriteLine("cleanup");}}
foreach(var _ in Gen()){} "#,
        ["cleanup"]
    };

    iterator_finally_prints_when_consumer_breaks_early => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{for(int i=0;i<5;i++)yield return i;}finally{Console.WriteLine("done");}}
foreach(var n in Gen()){if(n==2)break;}"#,
        ["done"]
    };

    iterator_finally_runs_once_per_enumeration => {
        r#"int fin=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin++;Console.WriteLine(fin);}}
foreach(var _ in Gen()){} foreach(var _ in Gen()){} "#,
        ["1", "2"]
    };

    yield_return_empty_sequence_count_zero => {
        r#"System.Collections.Generic.IEnumerable<int> Empty(){yield break;}
int c=0; foreach(var _ in Empty()) c++; Console.WriteLine(c);"#,
        ["0"]
    };

    yield_return_lazy_body_not_run_until_move_next => {
        r#"int calls=0; System.Collections.Generic.IEnumerable<int> Lazy(){calls++;yield return 1;}
var seq=Lazy(); Console.WriteLine(calls); foreach(var _ in seq){} Console.WriteLine(calls);"#,
        ["0", "1"]
    };

    yield_return_restart_iterator_second_foreach => {
        r#"System.Collections.Generic.IEnumerable<int> Two(){yield return 10;yield return 20;}
int sum=0; foreach(var n in Two()) sum+=n; foreach(var n in Two()) sum+=n; Console.WriteLine(sum);"#,
        ["60"]
    };

    yield_return_with_local_state_accumulator => {
        r#"System.Collections.Generic.IEnumerable<int> Running(){int s=0; for(int i=1;i<=3;i++){s+=i;yield return s;}}
Console.WriteLine(string.Join(",",Running()));"#,
        ["1,3,6"]
    };

    yield_return_filter_with_continue_skip => {
        r#"System.Collections.Generic.IEnumerable<int> Evens(int max){for(int i=0;i<=max;i++){if(i%2!=0)continue;yield return i;}}
Console.WriteLine(string.Join(",",Evens(6)));"#,
        ["0,2,4,6"]
    };

    yield_break_inside_nested_loop => {
        r#"System.Collections.Generic.IEnumerable<int> Grid(int rows,int cols){for(int r=0;r<rows;r++){for(int c=0;c<cols;c++){if(r==1&&c==1)yield break;yield return r*10+c;}}}
Console.WriteLine(string.Join(",",Grid(3,3)));"#,
        ["0,1,2,10,11"]
    };

    yield_return_string_chars_sequence => {
        r#"System.Collections.Generic.IEnumerable<char> Letters(string s){foreach(char c in s)yield return c;}
Console.WriteLine(string.Join("",Letters("ab")));"#,
        ["ab"]
    };

    yield_return_boxed_objects_via_enumerable => {
        r#"System.Collections.Generic.IEnumerable<object> Box(){yield return 1;yield return "two";}
int c=0; foreach(var _ in Box()) c++; Console.WriteLine(c);"#,
        ["2"]
    };

    nested_yield_return_with_inner_yield_break => {
        r#"System.Collections.Generic.IEnumerable<int> Inner(){yield return 1;yield break;yield return 9;}
System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in Inner())yield return x;yield return 2;}
Console.WriteLine(string.Join(",",Outer()));"#,
        ["1,2"]
    };

    iterator_try_finally_with_console_in_try_and_finally => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{Console.WriteLine("try");yield return 5;}finally{Console.WriteLine("finally");}}
foreach(var n in Gen()) Console.WriteLine(n);"#,
        ["try", "5", "finally"]
    };

    yield_return_from_local_function => {
        r#"System.Collections.Generic.IEnumerable<int> Outer(){System.Collections.Generic.IEnumerable<int> Inner(){yield return 3;} foreach(var x in Inner())yield return x;}
Console.WriteLine(Outer().First());"#,
        ["3"]
    };

    yield_return_multiple_enumerators_independent => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){yield return 1;yield return 2;}
var a=Gen(); var b=Gen(); Console.WriteLine(a.First()+b.First());"#,
        ["2"]
    };

    yield_return_dispose_pattern_prints_on_iterator_close => {
        r#"System.Collections.Generic.IEnumerable<int> Track(){try{yield return 1;}finally{Console.WriteLine("dispose");}}
using var e=Track().GetEnumerator(); e.MoveNext(); Console.WriteLine(e.Current);"#,
        ["1", "dispose"]
    };

    yield_return_in_do_while_emits_once => {
        r#"System.Collections.Generic.IEnumerable<int> Once(){int n=0;do{yield return n;n++;}while(n<1);}
Console.WriteLine(string.Join(",",Once()));"#,
        ["0"]
    };

    yield_return_with_nullable_int_values => {
        r#"System.Collections.Generic.IEnumerable<int?> Maybe(){yield return null;yield return 4;}
Console.WriteLine(string.Join(",",Maybe().Select(x=>x??0)));"#,
        ["0,4"]
    };

    yield_break_after_zero_yields_empty => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){yield break;yield return 1;}
Console.WriteLine(Gen().Count());"#,
        ["0"]
    };

    yield_return_nested_try_finally_inner_finally_print => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{try{yield return 1;}finally{Console.WriteLine("inner");}}finally{Console.WriteLine("outer");}}
foreach(var _ in Gen()){} "#,
        ["inner", "outer"]
    };

    yield_return_select_where_pipeline => {
        r#"System.Collections.Generic.IEnumerable<int> N(){for(int i=0;i<6;i++)yield return i;}
Console.WriteLine(N().Where(x=>x%2==0).Select(x=>x*10).Sum());"#,
        ["60"]
    };

    yield_return_enumerable_of_enumerable_flatten_manual => {
        r#"System.Collections.Generic.IEnumerable<System.Collections.Generic.IEnumerable<int>> Batches(){yield return new[]{1,2};yield return new[]{3};}
var flat=new System.Collections.Generic.List<int>(); foreach(var batch in Batches()) foreach(var n in batch) flat.Add(n); Console.WriteLine(string.Join(",",flat));"#,
        ["1,2,3"]
    };

    yield_return_with_parameterized_start_index => {
        r#"System.Collections.Generic.IEnumerable<int> From(int start,int count){for(int i=0;i<count;i++)yield return start+i;}
Console.WriteLine(string.Join(",",From(5,3)));"#,
        ["5,6,7"]
    };

    yield_return_break_on_condition_in_foreach_source => {
        r#"System.Collections.Generic.IEnumerable<int> TakeWhilePositive(int[] a){foreach(var n in a){if(n<0)yield break;yield return n;}}
Console.WriteLine(string.Join(",",TakeWhilePositive(new[]{2,4,-1,8})));"#,
        ["2,4"]
    };

    iterator_finally_not_run_until_started_iteration => {
        r#"int fin=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin=1;Console.WriteLine(fin);}}
var seq=Gen(); Console.WriteLine(fin); foreach(var _ in seq){} "#,
        ["0", "1"]
    };

    yield_return_infinite_prefix_take_three => {
        r#"System.Collections.Generic.IEnumerable<int> Naturals(){int n=0;while(true)yield return n++;}
Console.WriteLine(string.Join(",",Naturals().Take(3)));"#,
        ["0,1,2"]
    };

    nested_iterator_yields_concatenated_strings => {
        r#"System.Collections.Generic.IEnumerable<string> Words(){yield return "a";yield return "b";}
System.Collections.Generic.IEnumerable<string> Twice(){foreach(var w in Words())yield return w;foreach(var w in Words())yield return w;}
Console.WriteLine(string.Join("",Twice()));"#,
        ["a,b,a,b"]
    };

    yield_return_bool_sequence_all_true => {
        r#"System.Collections.Generic.IEnumerable<bool> Flags(){yield return true;yield return true;}
Console.WriteLine(Flags().All(x=>x));"#,
        ["True"]
    };

    yield_return_decimal_values_sum => {
        r#"System.Collections.Generic.IEnumerable<decimal> D(){yield return 1.5m;yield return 2.5m;}
Console.WriteLine(D().Sum());"#,
        ["4.0"]
    };

    iterator_finally_prints_even_if_no_yield_reached => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(bool ok){try{if(!ok)yield break;yield return 1;}finally{Console.WriteLine("end");}}
foreach(var _ in Gen(false)){} "#,
        ["end"]
    };

    yield_return_method_group_enumerated_by_foreach => {
        r#"System.Collections.Generic.IEnumerable<int> Range(int n){for(int i=0;i<n;i++)yield return i;}
void Run(System.Collections.Generic.IEnumerable<int> src){Console.WriteLine(src.Sum());}
Run(Range(5));"#,
        ["10"]
    };

    yield_return_with_explicit_ienumerable_interface => {
        r#"class Nums:System.Collections.Generic.IEnumerable<int>{public System.Collections.Generic.IEnumerator<int> GetEnumerator(){yield return 2;yield return 4;}System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();}
Console.WriteLine(new Nums().Sum());"#,
        ["6"]
    };

    nested_yield_with_conditional_inner_skip => {
        r#"System.Collections.Generic.IEnumerable<int> Inner(bool ok){if(ok)yield return 9;}
System.Collections.Generic.IEnumerable<int> Outer(bool ok){foreach(var x in Inner(ok))yield return x;yield return 1;}
Console.WriteLine(string.Join(",",Outer(false)));"#,
        ["1"]
    };

    yield_return_char_uppercase_projection => {
        r#"System.Collections.Generic.IEnumerable<char> Lower(){yield return 'a';yield return 'b';}
Console.WriteLine(string.Join(",",Lower().Select(c=>char.ToUpper(c))));"#,
        ["A,B"]
    };

    iterator_finally_runs_after_partial_to_list => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield return 2;yield return 3;}finally{Console.WriteLine("close");}}
Gen().Take(2).ToList(); "#,
        ["close"]
    };

    yield_break_at_start_before_any_yield => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){yield break;yield return 1;}
Console.WriteLine(Gen().Any());"#,
        ["False"]
    };

    yield_return_repeated_value_pattern => {
        r#"System.Collections.Generic.IEnumerable<int> Repeat(int v,int n){for(int i=0;i<n;i++)yield return v;}
Console.WriteLine(string.Join(",",Repeat(7,3)));"#,
        ["7,7,7"]
    };

    nested_iterator_count_matches_flat_length => {
        r#"System.Collections.Generic.IEnumerable<int> A(){yield return 1;yield return 2;}
System.Collections.Generic.IEnumerable<int> B(){foreach(var x in A())yield return x;}
Console.WriteLine(B().Count());"#,
        ["2"]
    };

    yield_return_with_struct_element_type => {
        r#"struct Pt{public int X;} System.Collections.Generic.IEnumerable<Pt> Points(){yield return new Pt{X=1};yield return new Pt{X=2};}
Console.WriteLine(Points().Sum(p=>p.X));"#,
        ["3"]
    };

    iterator_finally_print_order_after_last_element_read => {
        r#"System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 10;yield return 20;}finally{Console.WriteLine("after");}}
foreach(var n in Gen()) Console.WriteLine(n);"#,
        ["10", "20", "after"]
    };

    yield_return_from_generic_class_method => {
        r#"class Bag<T>{public System.Collections.Generic.IEnumerable<T> Single(T v){yield return v;}}
Console.WriteLine(new Bag<int>().Single(8).First());"#,
        ["8"]
    };

    yield_return_skip_take_window => {
        r#"System.Collections.Generic.IEnumerable<int> N(){for(int i=0;i<10;i++)yield return i;}
Console.WriteLine(string.Join(",",N().Skip(3).Take(2)));"#,
        ["3,4"]
    };

    nested_yield_return_with_outer_yield_break => {
        r#"System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in new[]{1,2,3}){if(x==2)yield break;yield return x;}}
Console.WriteLine(string.Join(",",Outer()));"#,
        ["1"]
    };

    iterator_disposal_finally_print_once_per_full_run => {
        r#"int hits=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{hits++;Console.WriteLine(hits);}}
foreach(var _ in Gen()){} Console.WriteLine(hits);"#,
        ["1", "1"]
    };

    yield_return_enumerable_as_return_type_of_helper => {
        r#"System.Collections.Generic.IEnumerable<int> Build(){yield return 3;yield return 5;}
int Total(){return Build().Sum();} Console.WriteLine(Total());"#,
        ["8"]
    };
}
