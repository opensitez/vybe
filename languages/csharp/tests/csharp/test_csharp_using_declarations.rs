//! `using var` declarations — scope-end disposal order and nested-scope LIFO cleanup.

csharp_cases! {
    using_var_disposes_after_following_statement => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var x=new R("x"); Console.WriteLine("body");"#,
        ["body", "x"]
    };

    using_var_disposes_before_outer_scope_ends => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
{using var x=new R("x"); Console.WriteLine("inner");} Console.WriteLine("outer");"#,
        ["inner", "x", "outer"]
    };

    using_var_two_in_same_scope_dispose_reverse_order => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var a=new R("a"); using var b=new R("b"); Console.WriteLine("done");"#,
        ["done", "b", "a"]
    };

    using_var_three_resources_lifo_on_scope_exit => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var one=new R("1"); using var two=new R("2"); using var three=new R("3"); Console.WriteLine("mid");"#,
        ["mid", "3", "2", "1"]
    };

    using_var_in_if_block_disposes_before_if_ends => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
if(true){using var x=new R("if"); Console.WriteLine("then");} Console.WriteLine("after");"#,
        ["then", "if", "after"]
    };

    using_var_in_else_block_disposes_on_else_exit => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
if(false){} else {using var x=new R("else"); Console.WriteLine("branch");} Console.WriteLine("end");"#,
        ["branch", "else", "end"]
    };

    using_var_in_while_loop_body_disposes_each_iteration => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int i=0; while(i<2){using var x=new R(i.ToString()); Console.WriteLine("loop"); i++;} Console.WriteLine("exit");"#,
        ["loop", "0", "loop", "1", "exit"]
    };

    using_var_in_for_loop_body_disposes_each_iteration => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
for(int i=0;i<2;i++){using var x=new R("f"+i); Console.WriteLine("iter");} Console.WriteLine("done");"#,
        ["iter", "f0", "iter", "f1", "done"]
    };

    using_var_in_foreach_body_disposes_per_element => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
foreach(var n in new[]{1,2}){using var x=new R("e"+n); Console.WriteLine("each");} Console.WriteLine("all");"#,
        ["each", "e1", "each", "e2", "all"]
    };

    using_var_nested_block_inner_then_outer_disposal => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var outer=new R("outer");
{using var inner=new R("inner"); Console.WriteLine("nest");}
Console.WriteLine("flat");"#,
        ["nest", "inner", "flat", "outer"]
    };

    using_var_before_return_from_local_function => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
string Read(){using var x=new R("fn"); return "ok";} Console.WriteLine(Read());"#,
        ["fn", "ok"]
    };

    using_var_with_early_return_still_disposes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int Go(bool stop){using var x=new R("go"); if(stop) return 1; return 2;} Console.WriteLine(Go(true));"#,
        ["go", "1"]
    };

    using_var_with_break_in_loop_still_disposes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
foreach(var n in new[]{1,2,3}){using var x=new R("b"); if(n==2) break; Console.WriteLine(n);} Console.WriteLine("end");"#,
        ["1", "b", "end"]
    };

    using_var_with_continue_in_loop_still_disposes_each_time => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
foreach(var n in new[]{1,2}){using var x=new R("c"); if(n==1) continue; Console.WriteLine(n);} Console.WriteLine("end");"#,
        ["c", "2", "c", "end"]
    };

    using_var_disposes_on_uncaught_exception_after_body_print => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
try{using var x=new R("boom"); Console.WriteLine("body"); throw new System.Exception();} catch{Console.WriteLine("caught");}"#,
        ["body", "boom", "caught"]
    };

    using_var_idisposable_interface_typed_variable => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using System.IDisposable x=new R("iface"); Console.WriteLine("use");"#,
        ["use", "iface"]
    };

    using_var_expression_bodied_dispose_prints_name => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;} public void Dispose()=>Console.WriteLine(n);}
using var x=new R("expr"); Console.WriteLine("ok");"#,
        ["ok", "expr"]
    };

    using_var_disposal_runs_after_all_prior_writes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var a=new R("last"); Console.WriteLine("1"); Console.WriteLine("2");"#,
        ["1", "2", "last"]
    };

    using_var_in_switch_case_block_disposes_on_case_exit => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
switch(1){case 1: using var x=new R("sw"); Console.WriteLine("case"); break;} Console.WriteLine("after");"#,
        ["case", "sw", "after"]
    };

    using_var_two_nested_blocks_dispose_inside_out => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
{using var a=new R("a"); {using var b=new R("b"); Console.WriteLine("deep");}} Console.WriteLine("shallow");"#,
        ["deep", "b", "a", "shallow"]
    };

    using_var_same_type_two_instances_reverse_dispose => {
        r#"class R:System.IDisposable{int id;public R(int id){this.id=id;}public void Dispose(){Console.WriteLine(id);}}
using var first=new R(1); using var second=new R(2); Console.WriteLine("pair");"#,
        ["pair", "2", "1"]
    };

    using_var_after_other_locals_still_lifo_with_peers => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int count=0; using var a=new R("a"); count++; using var b=new R("b"); count++; Console.WriteLine(count);"#,
        ["2", "b", "a"]
    };

    using_var_in_try_block_disposes_before_catch => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
try{using var x=new R("try"); throw new System.Exception();} catch{Console.WriteLine("catch");}"#,
        ["try", "catch"]
    };

    using_var_in_try_finally_finally_after_disposal => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
try{using var x=new R("res"); Console.WriteLine("try");} finally{Console.WriteLine("fin");}"#,
        ["try", "res", "fin"]
    };

    using_var_dispose_count_static_field_increments_once => {
        r#"class R:System.IDisposable{public static int N=0;public void Dispose(){N++; Console.WriteLine(N);}}
using var x=new R(); Console.WriteLine("once");"#,
        ["once", "1"]
    };

    using_var_two_scopes_sequential_dispose_runs_twice => {
        r#"class R:System.IDisposable{public static int N=0;public void Dispose(){N++;}}
{using var x=new R();} {using var y=new R();} Console.WriteLine(R.N);"#,
        ["2"]
    };

    using_var_memory_stream_length_ok_before_scope_end => {
        r#"using var ms=new System.IO.MemoryStream(new byte[]{1,2,3}); Console.WriteLine(ms.Length);"#,
        ["3"]
    };

    using_var_string_reader_reads_before_disposal => {
        r#"using var sr=new System.IO.StringReader("hi"); Console.WriteLine(sr.ReadLine());"#,
        ["hi"]
    };

    using_var_list_mutation_visible_before_dispose => {
        r#"using var list=new System.Collections.Generic.List<int>(); list.Add(4); Console.WriteLine(list[0]);"#,
        ["4"]
    };

    using_var_guard_clause_return_path_disposes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
string Pick(int v){using var x=new R("pick"); if(v<0) return "neg"; return "pos";} Console.WriteLine(Pick(-1));"#,
        ["pick", "neg"]
    };

    using_var_multiple_returns_same_scope_one_dispose => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int F(int n){using var x=new R("f"); if(n==0) return 0; if(n==1) return 1; return 2;} Console.WriteLine(F(1));"#,
        ["f", "1"]
    };

    using_var_in_local_function_nested_disposes_on_fn_exit => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
void Outer(){void Inner(){using var x=new R("in"); Console.WriteLine("fn");} Inner(); Console.WriteLine("out");} Outer();"#,
        ["fn", "in", "out"]
    };

    using_var_disposal_order_with_interleaved_console_writes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine("d:"+n);}}
using var a=new R("a"); Console.WriteLine("m1"); using var b=new R("b"); Console.WriteLine("m2");"#,
        ["m1", "m2", "d:b", "d:a"]
    };

    using_var_empty_block_still_disposes_immediately => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
{using var x=new R("empty");} Console.WriteLine("next");"#,
        ["empty", "next"]
    };

    using_var_in_conditional_expression_branch_not_taken_no_dispose => {
        r#"class R:System.IDisposable{public static int N=0;public void Dispose(){N++;}}
bool ok=true; if(ok){using var x=new R(); Console.WriteLine("yes");} else {using var y=new R();} Console.WriteLine(R.N);"#,
        ["yes", "1"]
    };

    using_var_lambda_invocation_disposes_after_lambda_returns => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
System.Func<int> f=()=>{using var x=new R("lam"); return 7;}; Console.WriteLine(f());"#,
        ["7", "lam"]
    };

    using_var_field_like_local_survives_until_scope_not_statement => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var hold=new R("hold"); for(int i=0;i<2;i++) Console.WriteLine(i);"#,
        ["0", "1", "hold"]
    };

    using_var_disposes_before_subsequent_using_var_peer_created_later => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
void Scope(){using var late=new R("late");} using var early=new R("early"); Scope(); Console.WriteLine("end");"#,
        ["late", "end", "early"]
    };

    using_var_in_do_while_executes_at_least_once_then_disposes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int n=0; do{using var x=new R("do"); Console.WriteLine("once"); n++;} while(n<1); Console.WriteLine("fin");"#,
        ["once", "do", "fin"]
    };

    using_var_with_nullable_reference_type_still_disposes => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var x=new R("nr"); Console.WriteLine(x==null?"null":"obj");"#,
        ["obj", "nr"]
    };

    using_var_chained_scope_three_levels_lifo => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
{using var l1=new R("l1"); {using var l2=new R("l2"); {using var l3=new R("l3"); Console.WriteLine("3");}} Console.WriteLine("2");} Console.WriteLine("1");"#,
        ["3", "l3", "2", "l2", "1", "l1"]
    };

    using_var_after_throw_in_same_block_disposes_before_propagation => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
try{using var x=new R("x"); throw new System.InvalidOperationException();} catch{Console.WriteLine("handled");}"#,
        ["x", "handled"]
    };

    using_var_disposal_prints_before_method_end_after_all_work => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
void Run(){using var x=new R("run"); Console.WriteLine("work");} Run(); Console.WriteLine("after");"#,
        ["work", "run", "after"]
    };

    using_var_two_methods_each_own_declaration_dispose_on_return => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
void A(){using var x=new R("A"); Console.WriteLine("a");}
void B(){using var x=new R("B"); Console.WriteLine("b");}
A(); B();"#,
        ["a", "A", "b", "B"]
    };

    using_var_in_temporary_block_between_statements => {
        r#"class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
Console.WriteLine("start"); {using var x=new R("mid"); Console.WriteLine("inside");} Console.WriteLine("finish");"#,
        ["start", "inside", "mid", "finish"]
    };
}
