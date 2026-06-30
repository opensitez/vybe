//! Finally return overrides try/catch return; break/continue from finally.

crate::js_cases! {
    finally_return_overrides_try_return => {
        r#"function f(){try{return "try";}finally{return "finally";}}console.log(f());"#,
        ["finally"]
    };

    finally_return_overrides_catch_return => {
        r#"function f(){try{throw 1;}catch{return "catch";}finally{return "finally";}}console.log(f());"#,
        ["finally"]
    };

    finally_return_after_caught_throw_in_try => {
        r#"function f(){try{throw new Error("t");}catch(e){return "c";}finally{return "f";}}console.log(f());"#,
        ["f"]
    };

    finally_without_return_preserves_try_value => {
        r#"function f(){try{return 42;}finally{}}console.log(f());"#,
        ["42"]
    };

    finally_throw_overrides_pending_try_return => {
        r#"function f(){try{return 1;}finally{throw new Error("fthrow");}}try{console.log(f());}catch(e){console.log(e.message);}"#,
        ["fthrow"]
    };

    nested_finally_inner_return_wins_over_outer => {
        r#"function f(){try{try{return "inner";}finally{return "fin";}}finally{return "outer";}}console.log(f());"#,
        ["fin"]
    };

    try_return_finally_logs_without_overriding => {
        r#"function f(){const o=[];try{return "ok";}finally{o.push("f");}return o.join(",");}console.log(f());"#,
        ["ok"]
    };

    catch_return_finally_overrides_to_later_value => {
        r#"function f(){try{throw 0;}catch{return "c";}finally{return "f";}}console.log(f());"#,
        ["f"]
    };

    loop_finally_break_exits_entire_loop => {
        r#"let o=[];for(let n=0;n<5;n++){try{o.push(n);if(n===2)throw n;}finally{if(n===2)break;}}console.log(o.join(","));"#,
        ["0,1,2"]
    };

    loop_finally_continue_skips_to_next_iteration => {
        r#"let o=[];for(let n=0;n<4;n++){try{o.push("t"+n);if(n===1)throw n;}catch{}finally{if(n===1)continue;}o.push("a"+n);}console.log(o.join(","));"#,
        ["t0,a0,t1,t2,a2,t3,a3"]
    };

    labeled_break_from_try_finally_in_loop => {
        r#"let o=[];outer:for(let i=0;i<3;i++){try{o.push(i);if(i===1)throw i;}finally{if(i===1)break outer;}}console.log(o.join(","));"#,
        ["0,1"]
    };

    finally_return_supersedes_throw_in_catch => {
        r#"function f(){try{throw 1;}catch{throw 2;}finally{return 9;}}console.log(f());"#,
        ["9"]
    };

    try_finally_no_catch_throw_reaches_caller => {
        r#"function f(){try{throw new Error("up");}finally{}}try{f();}catch(e){console.log(e.message);}"#,
        ["up"]
    };

    finally_runs_after_catch_rethrow => {
        r#"let o=[];try{try{throw 1;}catch(e){o.push("c");throw e;}}finally{o.push("f");}catch{o.push("o");}console.log(o.join(","));"#,
        ["c,f,o"]
    };

    finally_return_in_async_function => {
        r#"async function f(){try{return "a";}finally{return "b";}}f().then(v=>console.log(v));"#,
        ["b"]
    };

    nested_try_finally_return_chain => {
        r#"function f(){try{try{return 1;}finally{return 2;}}finally{return 3;}}console.log(f());"#,
        ["2"]
    };

    finally_with_break_inside_switch_from_try => {
        r#"let o=[];try{switch(1){case 1:try{throw "x";}finally{o.push("f");break;}default:o.push("d");}}catch(e){o.push(String(e));}console.log(o.join(","));"#,
        ["f,x"]
    };

    catch_finally_both_return_finally_wins => {
        r#"function f(){try{throw 0;}catch{try{return "c";}finally{return "cf";}}finally{return "ff";}}console.log(f());"#,
        ["cf"]
    };
}
