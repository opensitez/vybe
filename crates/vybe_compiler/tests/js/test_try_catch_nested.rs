//! Nested try/catch/finally — propagation, swallowing, try-in-catch, sibling blocks.

crate::js_cases! {
    nested_inner_swallows_thrown_error => {
        r#"let o=[];
try{try{throw new TypeError("x");}catch(e){o.push(e.name);}}
catch{o.push("outer");}
console.log(o.join(","));"#,
        ["TypeError"]
    };

    nested_inner_rethrows_error => {
        r#"let o=[];
try{try{throw new Error("a");}catch(e){o.push("inner");throw e;}}
catch(e){o.push("outer:"+e.message);}
console.log(o.join(","));"#,
        ["inner,outer:a"]
    };

    nested_triple_all_catch => {
        r#"let o=[];
try{try{try{throw new Error("d");}catch(e){o.push("L3");}}
catch(e){o.push("L2");}}
catch(e){o.push("L1");}
console.log(o.join(","));"#,
        ["L3"]
    };

    nested_triple_middle_rethrows => {
        // The middle arm RETHROWS (per the test's name/intent) so the outer
        // catch fires too (node-verified: "in,mid,out").
        r#"let o=[];
try{try{try{throw "x";}catch(e){o.push("in");throw e;}}
catch(e){o.push("mid");throw e;}}
catch(e){o.push("out");}
console.log(o.join(","));"#,
        ["in,mid,out"]
    };

    nested_triple_inner_finally => {
        r#"let o=[];
try{try{try{throw 1;}finally{o.push("f3");}}
catch(e){o.push("c2");}}
catch(e){o.push("c1");}
console.log(o.join(","));"#,
        ["f3,c2"]
    };

    nested_triple_outer_finally_only => {
        r#"let o=[];
try{try{throw 2;}catch(e){o.push("c");}}
finally{o.push("f");}
console.log(o.join(","));"#,
        ["c,f"]
    };

    nested_triple_no_throw => {
        r#"let o=[];
try{try{try{o.push("a");}catch{o.push("b");}}
catch{o.push("c");}}
catch{o.push("d");}
console.log(o.join(","));"#,
        ["a"]
    };

    nested_triple_finally_each_level => {
        r#"let o=[];
try{try{try{o.push("t");}finally{o.push("f1");}}
finally{o.push("f2");}}
finally{o.push("f3");}
console.log(o.join(","));"#,
        ["t,f1,f2,f3"]
    };

    nested_triple_throw_caught_at_level_two => {
        r#"let o=[];
try{try{try{throw new TypeError("t");}catch(e){throw e;}}
catch(e){o.push(e.name);}}
catch{o.push("outer");}
console.log(o.join(","));"#,
        ["TypeError"]
    };

    nested_triple_throw_caught_at_level_one => {
        r#"let o=[];
try{try{try{throw new RangeError("r");}catch(e){throw e;}}
catch(e){throw e;}}
catch(e){o.push(e.name);}
console.log(o.join(","));"#,
        ["RangeError"]
    };

    nested_triple_inner_return => {
        r#"function f(){let o=[];
try{try{try{return 1;}catch{return 2;}}
catch{return 3;}}
catch{return 4;}}
console.log(f());"#,
        ["1"]
    };

    nested_triple_catch_returns_value => {
        r#"function f(){let o=[];
try{try{throw 5;}catch(e){return "got:"+e;}}
catch{return "outer";}}
console.log(f());"#,
        ["got:5"]
    };

    nested_triple_finally_after_inner_catch => {
        r#"let o=[];
try{try{try{throw 0;}catch{o.push("c");}}
finally{o.push("f2");}}
finally{o.push("f1");}
console.log(o.join(","));"#,
        ["c,f2,f1"]
    };

    nested_triple_instanceof_filter_inner => {
        r#"let o=[];
try{try{try{throw new TypeError("x");}catch(e){if(e instanceof TypeError){o.push("ok");}else throw e;}}
catch{o.push("bad");}}
catch{o.push("worse");}
console.log(o.join(","));"#,
        ["ok"]
    };

    nested_triple_instanceof_filter_outer => {
        r#"let o=[];
try{try{try{throw new Error("x");}catch(e){throw e;}}
catch(e){if(e instanceof Error)o.push("e");else o.push("?");}}
catch{o.push("no");}
console.log(o.join(","));"#,
        ["e"]
    };

    nested_triple_string_throw_chain => {
        r#"let o=[];
try{try{try{throw "a";}catch(e){o.push(e);throw "b";}}
catch(e){o.push(e);throw "c";}}
catch(e){o.push(e);}
console.log(o.join(","));"#,
        ["a,b,c"]
    };

    nested_triple_null_throw_propagates => {
        r#"let o=[];
try{try{try{throw null;}catch(e){if(e===null){o.push("n");}else throw e;}}
catch{o.push("x");}}
catch{o.push("y");}
console.log(o.join(","));"#,
        ["n"]
    };

    try_inside_catch_recovers => {
        r#"let o=[];
try{throw new Error("fail");}
catch(e){try{o.push("caught");}catch{o.push("no");}}
console.log(o.join(","));"#,
        ["caught"]
    };

    try_inside_catch_nested_throw_caught => {
        r#"let o=[];
try{throw 1;}
catch(e){try{throw 2;}catch(x){o.push(x);}}
console.log(o.join(","));"#,
        ["2"]
    };

    try_inside_catch_finally_runs => {
        r#"let o=[];
try{throw "a";}
catch(e){try{o.push("t");}finally{o.push("f");}}
console.log(o.join(","));"#,
        ["t,f"]
    };

    try_inside_catch_return_from_inner => {
        // Top-level `return` is a SyntaxError (§16.1) — the concept needs a
        // function. The return from the inner try exits f, skipping the rest
        // of the catch body (node-verified: "inner").
        r#"let o=[];
function f(){ try{ throw 0; }catch(e){ try{ return "inner"; }catch{ return "no"; } o.push("skip"); } }
o.push(f());
console.log(o.join(","));"#,
        ["inner"]
    };

    try_inside_catch_outer_still_active => {
        r#"let o=[];
try{throw 1;}
catch(e){try{o.push("a");}catch{o.push("b");} o.push("c");}
console.log(o.join(","));"#,
        ["a,c"]
    };

    try_inside_catch_rethrow_from_inner => {
        r#"let o=[];
try{throw new TypeError("t");}
catch(e){try{throw new RangeError("r");}catch(x){o.push(x.name);}}
console.log(o.join(","));"#,
        ["RangeError"]
    };

    try_inside_catch_double_recovery => {
        r#"let o=[];
try{throw 1;}
catch(e){try{throw 2;}catch(x){try{o.push(x);}catch{o.push("z");}}}
console.log(o.join(","));"#,
        ["2"]
    };

    try_inside_catch_no_throw_in_inner => {
        r#"let o=[];
try{throw "x";}
catch(e){try{o.push("ok");}catch{o.push("bad");}}
console.log(o.join(","));"#,
        ["ok"]
    };

    try_inside_catch_finally_on_inner_throw => {
        r#"let o=[];
try{throw 0;}
catch(e){try{throw 1;}finally{o.push("fin");}}
console.log(o.join(","));"#,
        ["fin"]
    };

    try_inside_catch_optional_inner_catch => {
        r#"let o=[];
try{throw "e";}
catch(e){try{throw "i";} catch {o.push("opt");}}
console.log(o.join(","));"#,
        ["opt"]
    };

    try_inside_catch_preserves_outer_binding => {
        r#"let o=[];
try{throw {v:1};}
catch(e){try{o.push(e.v);}catch{o.push(0);}}
console.log(o.join(","));"#,
        ["1"]
    };

    try_inside_catch_with_inner_finally_return => {
        // Top-level `return` is a SyntaxError (§16.1) — wrapped in a
        // function; the finally runs on the way out (node-verified: "f").
        r#"let o=[];
function f(){ try{ throw 1; }catch(e){ try{ return "a"; }finally{ o.push("f"); } } }
f();
console.log(o.join(","));"#,
        ["f"]
    };

    try_inside_catch_sibling_after_inner => {
        r#"let o=[];
try{throw 1;}
catch(e){try{o.push("i");}catch{} o.push("s");}
console.log(o.join(","));"#,
        ["i,s"]
    };

    try_inside_catch_deep_three_inner => {
        r#"let o=[];
try{throw 0;}
catch(e){try{try{throw 1;}catch(x){o.push(x);}}catch{o.push("z");}}
console.log(o.join(","));"#,
        ["1"]
    };

    try_inside_catch_logs_before_inner => {
        r#"let o=[];
try{throw "out";}
catch(e){o.push("pre");try{o.push("in");}catch{}}
console.log(o.join(","));"#,
        ["pre,in"]
    };

    sibling_try_first_throws_second_runs => {
        r#"let o=[];
try{throw 1;}catch{o.push("a");}
try{o.push("b");}catch{o.push("c");}
console.log(o.join(","));"#,
        ["a,b"]
    };

    sibling_try_both_throw_caught_separately => {
        r#"let o=[];
try{throw "a";}catch(e){o.push(e);}
try{throw "b";}catch(e){o.push(e);}
console.log(o.join(","));"#,
        ["a,b"]
    };

    sibling_try_finally_independent => {
        r#"let o=[];
try{o.push("1");}finally{o.push("f1");}
try{o.push("2");}finally{o.push("f2");}
console.log(o.join(","));"#,
        ["1,f1,2,f2"]
    };

    sibling_try_first_finally_second_catch => {
        // Without a catch the first try's exception escapes and nothing
        // else runs — the intent (finally, then the sibling try) needs the
        // exception handled (node-verified: "f,ok").
        r#"let o=[];
try{throw 0;}catch{}finally{o.push("f");}
try{o.push("ok");}catch{o.push("c");}
console.log(o.join(","));"#,
        ["f,ok"]
    };

    sibling_nested_then_flat => {
        r#"let o=[];
try{try{o.push("n");}catch{}}catch{}
try{o.push("f");}catch{}
console.log(o.join(","));"#,
        ["n,f"]
    };

    sibling_three_sequential_catches => {
        r#"let o=[];
try{throw 1;}catch{o.push("1");}
try{throw 2;}catch{o.push("2");}
try{throw 3;}catch{o.push("3");}
console.log(o.join(","));"#,
        ["1,2,3"]
    };

    sibling_try_no_catch_then_with_catch => {
        r#"let o=[];
try{o.push("a");}finally{o.push("b");}
try{throw 1;}catch{o.push("c");}
console.log(o.join(","));"#,
        ["a,b,c"]
    };

    sibling_outer_wraps_two_inner => {
        r#"let o=[];
try{
  try{o.push("a");}catch{}
  try{throw 1;}catch{o.push("b");}
}catch{o.push("c");}
console.log(o.join(","));"#,
        ["a,b"]
    };

    sibling_finally_between_tries => {
        r#"let o=[];
try{o.push("t1");}finally{o.push("f");}
try{throw 0;}catch{o.push("c");}
console.log(o.join(","));"#,
        ["t1,f,c"]
    };

    sibling_independent_rethrow => {
        r#"let o=[];
try{try{throw "x";}catch(e){o.push("i");throw e;}}catch(e){o.push("o");}
try{o.push("next");}catch{o.push("nc");}
console.log(o.join(","));"#,
        ["i,o,next"]
    };

    nested_inner_finally_before_outer_catch => {
        r#"let o=[];
try{try{throw 1;}finally{o.push("f");}}
catch(e){o.push("c");}
console.log(o.join(","));"#,
        ["f,c"]
    };

    nested_outer_finally_after_inner_catch => {
        r#"let o=[];
try{try{throw 1;}catch{o.push("c");}}
finally{o.push("f");}
console.log(o.join(","));"#,
        ["c,f"]
    };

    nested_both_finally_inner_first => {
        r#"let o=[];
try{try{o.push("t");}finally{o.push("fi");}}
finally{o.push("fo");}
console.log(o.join(","));"#,
        ["t,fi,fo"]
    };

    // Node-verified: the original put `catch` AFTER `finally` — a
    // SyntaxError (§14.15). Valid form: outer try/catch around the
    // try/finally that wraps the rethrowing catch.
    nested_inner_finally_on_rethrow => {
        r#"let o=[];
try{
  try{
    try{ throw 1; }catch(e){ o.push("c"); throw e; }
  }finally{ o.push("f"); }
}catch(e){ o.push("o"); }
console.log(o.join(","));"#,
        ["c,f,o"]
    };

    nested_catch_in_inner_finally_in_outer => {
        r#"let o=[];
try{try{throw "x";}catch(e){o.push(e);}}
finally{o.push("f");}
console.log(o.join(","));"#,
        ["x,f"]
    };

    nested_no_throw_both_finally => {
        r#"let o=[];
try{try{o.push("a");}finally{o.push("b");}}
finally{o.push("c");}
console.log(o.join(","));"#,
        ["a,b,c"]
    };

    nested_inner_try_finally_no_catch => {
        r#"let o=[];
try{try{throw 0;}finally{o.push("f");}}
catch{o.push("c");}
console.log(o.join(","));"#,
        ["f,c"]
    };

    nested_outer_catch_skipped_inner_swallows => {
        r#"let o=[];
try{try{throw new EvalError("e");}catch(e){o.push("s");}}
catch{o.push("o");}
console.log(o.join(","));"#,
        ["s"]
    };

    nested_error_object_identity_preserved => {
        r#"let obj={id:7};
let same=false;
try{try{throw obj;}catch(e){same=(e===obj);}}
catch{same=false;}
console.log(same);"#,
        ["true"]
    };

    nested_number_throw_caught_as_number => {
        r#"let n=0;
try{try{throw 42;}catch(e){n=e;}}
catch{n=-1;}
console.log(n);"#,
        ["42"]
    };

    nested_inner_catch_mutates_and_rethrows => {
        r#"let o=[];
try{try{throw {m:"a"};}catch(e){e.m="b";throw e;}}
catch(e){o.push(e.m);}
console.log(o.join(","));"#,
        ["b"]
    };

    nested_async_style_sync_nested => {
        r#"let o=[];
try{try{throw "sync";}catch(e){o.push("1:"+e);}}
catch(e){o.push("2:"+e);}
console.log(o.join(","));"#,
        ["1:sync"]
    };

    nested_inner_optional_catch_outer_named => {
        r#"let o=[];
try{try{throw 1;}catch{o.push("i");}}
catch(e){o.push("o:"+e);}
console.log(o.join(","));"#,
        ["i"]
    };

    nested_outer_optional_inner_named => {
        r#"let o=[];
try{try{throw 2;}catch(e){o.push("i:"+e);}}
catch{o.push("o");}
console.log(o.join(","));"#,
        ["i:2"]
    };

    nested_finally_logs_on_no_error => {
        r#"let o=[];
try{try{o.push("ok");}finally{o.push("f");}}
catch{o.push("c");}
console.log(o.join(","));"#,
        ["ok,f"]
    };

}
