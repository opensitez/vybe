//! Throw inside loops — for, while, do-while, for-of, for-in, labeled break/continue.

crate::js_cases! {
    for_loop_throw_stops_at_first_caught_iteration => {
        r#"let o=[];for(let i=0;i<4;i++){try{if(i===2)throw new Error("mid");o.push(i);}catch(e){o.push("c");break;}}console.log(o.join(","));"#,
        ["0,1,c"]
    };

    for_loop_throw_uncaught_propagates_out => {
        r#"let o=[];try{for(let i=0;i<2;i++){if(i===1)throw new Error("out");o.push(i);}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["0,out"]
    };

    for_loop_continue_after_caught_throw_skips_body => {
        r#"let o=[];for(let i=0;i<3;i++){try{if(i===1)throw "skip";o.push(i);}catch{o.push("x");}}console.log(o.join(","));"#,
        ["0,x,2"]
    };

    while_loop_throw_on_condition => {
        r#"let o=[],n=0;while(n<3){try{if(n===1)throw new RangeError("w");o.push(n);}catch(e){o.push(e.name);}n++;}console.log(o.join(","));"#,
        ["0,RangeError,2"]
    };

    do_while_executes_body_before_throw_check => {
        r#"let o=[],n=0;do{try{o.push(n);if(n===0)throw "once";}catch(e){o.push(String(e));}n++;}while(n<2);console.log(o.join(","));"#,
        ["0,once,1"]
    };

    for_of_throw_from_iterator_value => {
        r#"let o=[];try{for(const x of [1,2,3]){if(x===2)throw new Error("v"+x);o.push(x);}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["1,v2"]
    };

    for_in_throw_custom_object_key => {
        r#"let o=[];const obj={a:1,b:2};try{for(const k in obj){if(k==="b")throw new Error(k);o.push(k);}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["a,b"]
    };

    labeled_break_from_try_inside_loop => {
        r#"let o=[];outer:for(let i=0;i<3;i++){try{if(i===1)throw "stop";o.push(i);}catch(e){o.push("b");break outer;}}console.log(o.join(","));"#,
        ["0,b"]
    };

    labeled_continue_from_catch_resumes_loop => {
        r#"let o=[];outer:for(let i=0;i<3;i++){try{if(i===1)throw "x";o.push(i);}catch{continue outer;}o.push("a");}console.log(o.join(","));"#,
        ["0,a,2,a"]
    };

    nested_loop_inner_throw_caught_by_outer_try => {
        r#"let o=[];try{for(let i=0;i<2;i++){for(let j=0;j<2;j++){if(i===1&&j===1)throw new Error("deep");o.push(i+""+j);}}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["00,01,10,deep"]
    };

    for_loop_throw_in_update_expression => {
        r#"let o=[];try{for(let i=0;i<3;(()=>{throw new Error("upd")})()){o.push(i);}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["0,upd"]
    };

    for_loop_throw_in_init_expression => {
        r#"let o=[];try{for(let i=(()=>{throw new Error("init");})();i<1;i++)o.push(i);}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["init"]
    };

    while_true_break_from_catch_after_throw => {
        r#"let o=[],n=0;while(true){try{if(n===2)throw "done";o.push(n);}catch{break;}n++;}console.log(o.join(","));"#,
        ["0,1"]
    };

    for_of_throw_null_value => {
        r#"let o=[];try{for(const x of [1,null,3]){if(x===null)throw new TypeError("null");o.push(x);}}catch(e){o.push(e.name);}console.log(o.join(","));"#,
        ["1,TypeError"]
    };

    for_in_throw_when_enumerable_key_is_symbol_skipped => {
        r#"let o=[];const s=Symbol("s");const obj={a:1};obj[s]=2;try{for(const k in obj){if(k==="a")throw new Error("a");o.push(k);}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["a"]
    };

    loop_try_finally_runs_after_throw => {
        r#"let o=[];for(let i=0;i<2;i++){try{try{if(i===1)throw "t";o.push(i);}finally{o.push("f");}}catch(e){o.push(String(e));}}console.log(o.join(","));"#,
        ["0,f,f,t"]
    };

    for_await_throw_from_async_iterator => {
        r#"async function main(){const o=[];async function* g(){yield 1;throw new Error("async");}try{for await(const v of g())o.push(v);}catch(e){o.push(e.message);}console.log(o.join(","));}main();"#,
        ["1,async"]
    };

    for_loop_rethrow_changes_caught_value => {
        r#"let o=[];for(let i=0;i<2;i++){try{if(i===1)throw "a";}catch(e){try{throw "b";}catch(x){o.push(x);}}}console.log(o.join(","));"#,
        ["b"]
    };

    do_while_false_condition_still_throws_once => {
        r#"let o=[];let run=true;do{try{if(run)throw new Error("once");}catch(e){o.push(e.message);run=false;}o.push("body");}while(false);console.log(o.join(","));"#,
        ["once,body"]
    };

    for_loop_empty_body_throw_in_condition => {
        r#"let o=[];try{for(let i=0;i<3&&(()=>{if(i===2)throw new Error("cond");return true;})();i++){o.push(i);}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["0,1,cond"]
    };

    infinite_while_caught_throw_counter => {
        r#"let o=[],n=0;while(n<5){try{if(n===3)throw n;}catch(e){o.push("c"+e);}o.push(n);n++;}console.log(o.join(","));"#,
        ["0,1,2,c3,3,4"]
    };

    for_of_break_from_try_before_throw => {
        r#"let o=[];for(const x of [1,2,3]){try{if(x===2)break;o.push(x);}catch{o.push("e");}}console.log(o.join(","));"#,
        ["1"]
    };
}
