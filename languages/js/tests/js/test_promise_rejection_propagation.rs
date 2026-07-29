//! Promise rejection propagation — .then throw, .catch recovery, chaining,
//! Promise.resolve/reject with thenables that throw.

crate::js_cases! {
    then_throw_string_reason_reaches_catch => {
        r#"Promise.resolve(1).then(()=>{throw "boom";}).catch(e=>console.log(e)).then(()=>console.log("ok"));"#,
        ["boom", "ok"]
    };

    then_throw_typeerror_preserves_name => {
        r#"Promise.resolve(0).then(()=>{throw new TypeError("bad");}).catch(e=>console.log(e.name));"#,
        ["TypeError"]
    };

    then_throw_rangeerror_preserves_message => {
        r#"Promise.resolve(0).then(()=>{throw new RangeError("out");}).catch(e=>console.log(e.message));"#,
        ["out"]
    };

    then_throw_reference_error_propagates => {
        r#"Promise.resolve(0).then(()=>{throw new ReferenceError("undef");}).catch(e=>console.log(e instanceof ReferenceError));"#,
        ["true"]
    };

    then_throw_syntax_error_propagates => {
        r#"Promise.resolve(0).then(()=>{throw new SyntaxError("parse");}).catch(e=>console.log(e.name));"#,
        ["SyntaxError"]
    };

    then_throw_undefined_reason => {
        r#"Promise.resolve(1).then(()=>{throw undefined;}).catch(e=>console.log(e===undefined));"#,
        ["true"]
    };

    then_throw_null_reason => {
        r#"Promise.resolve(1).then(()=>{throw null;}).catch(e=>console.log(e===null));"#,
        ["true"]
    };

    then_throw_number_reason => {
        r#"Promise.resolve(1).then(()=>{throw 404;}).catch(e=>console.log(e));"#,
        ["404"]
    };

    then_throw_plain_object_reason => {
        r#"Promise.resolve(1).then(()=>{throw {code:"E"};}).catch(e=>console.log(e.code));"#,
        ["E"]
    };

    then_throw_after_value_transform => {
        r#"Promise.resolve(2).then(x=>x*3).then(()=>{throw "late";}).catch(e=>console.log(e));"#,
        ["late"]
    };

    then_throw_in_second_handler_skips_third => {
        r#"const o=[];Promise.resolve(1).then(v=>{o.push("a");return v;}).then(()=>{throw 1;}).then(()=>o.push("b")).catch(()=>o.push("c")).then(()=>console.log(o.join(",")));"#,
        ["a,c"]
    };

    then_rejection_skips_fulfilled_handler => {
        r#"const o=[];Promise.reject("x").then(()=>o.push("then")).catch(()=>o.push("catch")).then(()=>console.log(o.join(",")));"#,
        ["catch"]
    };

    then_return_rejected_promise_propagates => {
        r#"Promise.resolve(1).then(()=>Promise.reject("r")).catch(e=>console.log(e));"#,
        ["r"]
    };

    then_return_rejected_promise_vs_throw_same_outcome => {
        r#"const o=[];Promise.resolve(0).then(()=>Promise.reject("a")).catch(e=>o.push("1:"+e));Promise.resolve(0).then(()=>{throw "a";}).catch(e=>o.push("2:"+e)).then(()=>console.log(o.join("|")));"#,
        ["1:a", "2:a"]
    };

    then_second_arg_handles_rejection_without_catch => {
        r#"Promise.reject("no").then(v=>console.log("ful:"+v),e=>console.log("rej:"+e));"#,
        ["rej:no"]
    };

    then_second_arg_skips_trailing_catch => {
        r#"const o=[];Promise.reject("x").then(v=>o.push("f:"+v),e=>o.push("h:"+e)).catch(e=>o.push("c:"+e)).then(()=>console.log(o.join(",")));"#,
        ["h:x"]
    };

    then_second_arg_not_called_on_fulfillment => {
        r#"Promise.resolve(5).then(v=>console.log(v),e=>console.log("rej:"+e));"#,
        ["5"]
    };

    catch_returns_scalar_recovers_chain => {
        r#"Promise.reject("err").catch(()=>"fixed").then(v=>console.log(v));"#,
        ["fixed"]
    };

    catch_returns_promise_flattens_value => {
        r#"Promise.reject("e").catch(()=>Promise.resolve(9)).then(v=>console.log(v));"#,
        ["9"]
    };

    catch_returns_rejected_promise_stays_rejected => {
        r#"Promise.reject("a").catch(()=>Promise.reject("b")).catch(e=>console.log(e));"#,
        ["b"]
    };

    catch_throws_new_error_replaces_reason => {
        r#"Promise.reject("old").catch(()=>{throw new Error("new");}).catch(e=>console.log(e.message));"#,
        ["new"]
    };

    catch_recovery_then_transforms_value => {
        r#"Promise.reject(1).catch(e=>e+1).then(v=>console.log(v*2));"#,
        ["4"]
    };

    catch_only_handles_nearest_rejection => {
        r#"Promise.reject("x").catch(e=>"got:"+e).then(v=>console.log(v));"#,
        ["got:x"]
    };

    catch_after_multiple_thens_receives_original => {
        r#"Promise.resolve(1).then(x=>x+1).then(x=>x*2).then(()=>{throw "mid";}).catch(e=>console.log(e));"#,
        ["mid"]
    };

    catch_handler_receives_original_reason_object => {
        r#"const r={id:3};Promise.reject(r).catch(e=>console.log(e===r));"#,
        ["true"]
    };

    chain_throw_catch_throw_catch => {
        r#"Promise.resolve(0).then(()=>{throw "a";}).catch(()=>{throw "b";}).catch(e=>console.log(e));"#,
        ["b"]
    };

    chain_reject_catch_then_throw_catch => {
        r#"Promise.reject("r").catch(e=>"ok:"+e).then(v=>{throw v+"!";}).catch(e=>console.log(e));"#,
        ["ok:r!"]
    };

    chain_multiple_catch_only_first_fires => {
        r#"const o=[];Promise.reject(1).catch(e=>o.push("c1:"+e)).catch(e=>o.push("c2:"+e)).then(()=>console.log(o.join(",")));"#,
        ["c1:1"]
    };

    chain_catch_then_throw_reaches_outer_catch => {
        r#"Promise.reject(0).catch(()=>1).then(()=>{throw 2;}).catch(e=>console.log(e));"#,
        ["2"]
    };

    promise_reject_string_reason => {
        r#"Promise.reject("denied").catch(e=>console.log(typeof e));"#,
        ["string"]
    };

    promise_reject_error_preserves_message => {
        r#"Promise.reject(new Error("fail")).catch(e=>console.log(e.message));"#,
        ["fail"]
    };

    promise_reject_typeerror_instanceof => {
        r#"Promise.reject(new TypeError("t")).catch(e=>console.log(e instanceof TypeError));"#,
        ["true"]
    };

    promise_reject_without_handler_caught_at_end => {
        r#"Promise.reject("late").then(()=>{}).then(()=>{}).catch(e=>console.log(e));"#,
        ["late"]
    };

    promise_resolve_thenable_then_throws_sync => {
        r#"Promise.resolve({then(res){res(1);}}).then(()=>{throw "sync";}).catch(e=>console.log(e));"#,
        ["sync"]
    };

    promise_resolve_thenable_then_throws_after_resolve => {
        r#"Promise.resolve({then(res){res(5);}}).then(v=>v).then(()=>{throw "after";}).catch(e=>console.log(e));"#,
        ["after"]
    };

    promise_resolve_thenable_rejects_then_catch_recovers => {
        r#"Promise.resolve({then(_,rej){rej("fromThenable");}}).catch(e=>"caught:"+e).then(v=>console.log(v));"#,
        ["caught:fromThenable"]
    };

    promise_resolve_thenable_throws_in_then_method => {
        r#"Promise.resolve({then(){throw new Error("thenThrow");}}).catch(e=>console.log(e.message));"#,
        ["thenThrow"]
    };

    promise_resolve_thenable_returns_rejecting_promise => {
        r#"Promise.resolve({then(res){res(Promise.reject("inner"));}}).catch(e=>console.log(e));"#,
        ["inner"]
    };

    promise_resolve_nested_thenable_chain => {
        r#"Promise.resolve({then(res){res({then(r2){r2(7);}});}}).then(v=>console.log(v));"#,
        ["7"]
    };

    promise_resolve_thenable_throw_before_calling_res => {
        r#"Promise.resolve({then(){throw "early";}}).catch(e=>console.log(e));"#,
        ["early"]
    };

    promise_resolve_non_thenable_wraps_value => {
        r#"Promise.resolve(42).then(v=>{throw v;}).catch(e=>console.log(e));"#,
        ["42"]
    };

    promise_resolve_promise_identity_adopts_state => {
        r#"Promise.resolve(Promise.reject("adopted")).catch(e=>console.log(e));"#,
        ["adopted"]
    };

    promise_reject_thenable_stays_rejected_reason => {
        r#"const t={then(){}};Promise.reject(t).catch(e=>console.log(e===t));"#,
        ["true"]
    };

    then_on_rejected_promise_with_throw_in_fulfill => {
        r#"Promise.reject("skip").then(()=>{throw "never";}).catch(e=>console.log(e));"#,
        ["skip"]
    };

    then_handler_throw_preserves_stack_as_error => {
        r#"Promise.resolve(0).then(()=>{throw new Error("stack");}).catch(e=>console.log(e instanceof Error));"#,
        ["true"]
    };

    catch_without_rethrow_continues_as_resolved => {
        r#"Promise.reject("x").catch(()=>{}).then(()=>console.log("continued"));"#,
        ["continued"]
    };

    catch_return_undefined_resolves_undefined => {
        r#"Promise.reject("x").catch(()=>{}).then(v=>console.log(v===undefined));"#,
        ["true"]
    };

    then_after_catch_receives_recovery_value => {
        r#"Promise.reject("e").catch(()=>"rec").then(v=>console.log("v:"+v));"#,
        ["v:rec"]
    };

    long_chain_single_throw_at_end => {
        r#"Promise.resolve(1).then(x=>x+1).then(x=>x*2).then(x=>x-1).then(()=>{throw "end";}).catch(e=>console.log(e));"#,
        ["end"]
    };

    long_chain_throw_in_middle_catch_at_end => {
        r#"Promise.resolve(1).then(x=>x+1).then(()=>{throw "mid";}).then(x=>x*10).catch(e=>console.log(e));"#,
        ["mid"]
    };

    then_throw_boolean_reason => {
        r#"Promise.resolve(0).then(()=>{throw false;}).catch(e=>console.log(String(e)));"#,
        ["false"]
    };

    then_throw_bigint_reason => {
        r#"Promise.resolve(0).then(()=>{throw 99n;}).catch(e=>console.log(String(e)));"#,
        ["99"]
    };

    then_throw_symbol_reason => {
        r#"const s=Symbol("s");Promise.resolve(0).then(()=>{throw s;}).catch(e=>console.log(e===s));"#,
        ["true"]
    };

    then_throw_array_reason => {
        r#"Promise.resolve(0).then(()=>{throw [1,2];}).catch(e=>console.log(e.join(",")));"#,
        ["1,2"]
    };

    catch_rethrows_same_reason => {
        r#"Promise.reject("same").catch(e=>{throw e;}).catch(e=>console.log(e));"#,
        ["same"]
    };

    catch_wraps_reason_in_new_error => {
        r#"Promise.reject("raw").catch(e=>{throw new Error("wrap:"+e);}).catch(e=>console.log(e.message));"#,
        ["wrap:raw"]
    };

    then_fulfilled_handler_not_called_after_throw => {
        r#"const o=[];Promise.resolve(1).then(()=>{throw 1;}).then(()=>o.push("nope")).catch(()=>o.push("yes")).then(()=>console.log(o.join(",")));"#,
        ["yes"]
    };

    parallel_rejections_independent_catches => {
        r#"const o=[];Promise.reject("a").catch(e=>o.push(e));Promise.reject("b").catch(e=>o.push(e));Promise.resolve().then(()=>console.log(o.sort().join(",")));"#,
        ["a,b"]
    };

    promise_constructor_executor_throw_becomes_rejection => {
        r#"new Promise(()=>{throw "exec";}).catch(e=>console.log(e));"#,
        ["exec"]
    };

    promise_constructor_reject_then_throw_in_then => {
        r#"new Promise((_,rej)=>rej("init")).then(()=>{throw "nope";}).catch(e=>console.log(e));"#,
        ["init"]
    };

    thenable_with_both_then_and_catch_like_methods => {
        r#"Promise.resolve({then(res,rej){res(1);}}).then(()=>{throw "x";}).catch(e=>console.log(e));"#,
        ["x"]
    };

    thenable_resolve_then_reject_in_same_then => {
        r#"Promise.resolve({then(res,rej){res(1);rej(2);}}).then(v=>console.log("v:"+v));"#,
        ["v:1"]
    };

    thenable_async_resolve_then_throw => {
        r#"Promise.resolve({then(res){queueMicrotask(()=>res(3));}}).then(()=>{throw "async";}).catch(e=>console.log(e));"#,
        ["async"]
    };

    catch_on_fulfilled_promise_never_runs => {
        r#"const o=[];Promise.resolve(1).catch(()=>o.push("catch")).then(()=>o.push("then")).then(()=>console.log(o.join(",")));"#,
        ["then"]
    };

    then_throw_inside_conditional => {
        r#"Promise.resolve(true).then(ok=>{if(ok)throw "cond";}).catch(e=>console.log(e));"#,
        ["cond"]
    };

    then_throw_inside_loop => {
        r#"Promise.resolve(3).then(n=>{for(let i=0;i<n;i++){if(i===2)throw "loop";}}).catch(e=>console.log(e));"#,
        ["loop"]
    };

    catch_modifies_reason_before_rethrow => {
        r#"Promise.reject({m:"a"}).catch(e=>{e.m="b";throw e;}).catch(e=>console.log(e.m));"#,
        ["b"]
    };

    chain_alternating_resolve_reject_with_catch => {
        r#"Promise.resolve(1).then(()=>Promise.reject("r")).catch(e=>e+"!").then(v=>console.log(v));"#,
        ["r!"]
    };

    then_second_arg_returns_recovery_value => {
        r#"Promise.reject("x").then(null,e=>"rec:"+e).then(v=>console.log(v));"#,
        ["rec:x"]
    };

    then_second_arg_throws_propagates => {
        r#"Promise.reject("x").then(null,()=>{throw "y";}).catch(e=>console.log(e));"#,
        ["y"]
    };

    then_null_fulfilled_skips_to_catch_on_reject => {
        r#"Promise.reject("z").then(null).catch(e=>console.log(e));"#,
        ["z"]
    };

    nested_promise_in_then_throw_bubbles => {
        r#"Promise.resolve(0).then(()=>Promise.resolve().then(()=>{throw "nested";})).catch(e=>console.log(e));"#,
        ["nested"]
    };

    rejection_from_awaited_promise_in_then => {
        r#"Promise.resolve(0).then(async()=>{throw "asyncThrow";}).catch(e=>console.log(e));"#,
        ["asyncThrow"]
    };

    catch_on_already_resolved_after_throw_in_sibling => {
        r#"const o=[];const p=Promise.resolve(1);p.then(()=>{throw "a";}).catch(e=>o.push(e));p.then(v=>o.push("sib:"+v)).then(()=>console.log(o.join("|")));"#,
        ["a", "sib:1"]
    };

    promise_all_rejection_propagates_to_catch => {
        r#"Promise.all([Promise.resolve(1),Promise.reject("all")]).catch(e=>console.log(e));"#,
        ["all"]
    };

    promise_race_rejection_reaches_catch => {
        r#"Promise.race([new Promise((_,r)=>setTimeout(()=>r("slow"),50)),Promise.reject("fast")]).catch(e=>console.log(e));"#,
        ["fast"]
    };

    then_catch_then_preserves_order => {
        r#"const o=[];Promise.reject(0).catch(()=>o.push("c")).then(()=>o.push("t")).then(()=>console.log(o.join(",")));"#,
        ["c,t"]
    };

    throw_custom_error_subclass => {
        r#"class MyErr extends Error{}Promise.resolve(0).then(()=>{throw new MyErr("custom");}).catch(e=>console.log(e instanceof MyErr));"#,
        ["true"]
    };

    catch_does_not_see_fulfilled_throw_from_later_then => {
        r#"Promise.resolve(1).then(()=>1).then(()=>{throw "later";}).catch(e=>console.log(e));"#,
        ["later"]
    };

    promise_resolve_thenable_with_getter_then => {
        r#"const o={get then(){throw "getter";}};Promise.resolve(o).catch(e=>console.log(e));"#,
        ["getter"]
    };

    thenable_calling_reject_after_resolve_ignored => {
        r#"Promise.resolve({then(res,rej){res(1);rej(2);}}).then(v=>console.log(v));"#,
        ["1"]
    };

    rejection_reason_survives_multiple_thens => {
        r#"Promise.reject("persist").then(null,e=>e+"!").then(v=>console.log(v));"#,
        ["persist!"]
    };

    catch_return_promise_that_throws_in_then => {
        r#"Promise.reject("a").catch(()=>Promise.resolve().then(()=>{throw "b";})).catch(e=>console.log(e));"#,
        ["b"]
    };

    then_immediately_after_catch_receives_value => {
        r#"Promise.reject("e").catch(()=>"ok").then(v=>console.log(v+"!"));"#,
        ["ok!"]
    };

    throw_in_then_with_onrejected_handler_bypasses_it => {
        r#"Promise.resolve(1).then(()=>{throw "t";},()=>console.log("skip")).catch(e=>console.log(e));"#,
        ["t"]
    };

    promise_reject_caught_by_nearest_catch_only => {
        r#"const o=[];Promise.reject(1).catch(e=>o.push("n:"+e)).then(()=>o.push("between")).catch(e=>o.push("far:"+e)).then(()=>console.log(o.join(",")));"#,
        ["n:1,between"]
    };

    then_return_throw_expression_short_circuit => {
        r#"Promise.resolve(0).then(()=>{const x=()=>{throw "fn";};return x();}).catch(e=>console.log(e));"#,
        ["fn"]
    };

    catch_logs_and_returns_for_chain_continue => {
        r#"const o=[];Promise.reject("x").catch(e=>{o.push("log:"+e);return "go";}).then(v=>o.push(v)).then(()=>console.log(o.join(",")));"#,
        ["log:x,go"]
    };

    promise_resolve_thenable_with_throw_in_microtask => {
        r#"Promise.resolve({then(res){queueMicrotask(()=>{throw "mt";});}}).catch(e=>console.log(e));"#,
        ["mt"]
    };

    then_handler_receives_undefined_from_void_return => {
        r#"Promise.resolve(5).then(()=>{}).then(v=>{throw "u:"+v;}).catch(e=>console.log(e));"#,
        ["u:undefined"]
    };

    multiple_sequential_catches_on_same_promise_branch => {
        r#"Promise.reject("a").catch(e=>"b:"+e).then(v=>{throw v;}).catch(e=>console.log(e));"#,
        ["b:a"]
    };

    rejection_from_promise_in_then_return => {
        r#"Promise.resolve(0).then(()=>new Promise((_,r)=>r("ctor"))).catch(e=>console.log(e));"#,
        ["ctor"]
    };

    then_throw_after_catch_recovery_restarts_chain => {
        r#"Promise.reject("a").catch(()=>"ok").then(v=>{throw v+"!";}).catch(e=>console.log(e));"#,
        ["ok!"]
    };

    promise_resolve_thenable_that_is_promise_reject => {
        r#"Promise.resolve(Promise.reject("deep")).catch(e=>console.log(e));"#,
        ["deep"]
    };

    catch_on_sync_throw_equivalent_to_reject => {
        r#"const o=[];Promise.resolve().then(()=>{throw "s";}).catch(e=>o.push("s:"+e));Promise.reject("a").catch(e=>o.push("a:"+e)).then(()=>console.log(o.sort().join("|")));"#,
        ["a:a", "s:s"]
    };

    then_rejection_handler_called_with_undefined_fulfill => {
        r#"Promise.reject("r").then(undefined,e=>console.log("h:"+e));"#,
        ["h:r"]
    };

    then_fulfillment_handler_undefined_skips_to_catch => {
        r#"Promise.resolve(9).then(undefined).then(v=>console.log(v));"#,
        ["9"]
    };

    chain_catch_at_beginning_not_possible_use_then_second => {
        r#"Promise.reject("x").then(null,e=>"handled").then(v=>console.log(v));"#,
        ["handled"]
    };

    throw_error_with_cause_property => {
        r#"const c=new Error("cause");Promise.resolve(0).then(()=>{const e=new Error("main");e.cause=c;throw e;}).catch(e=>console.log(e.cause.message));"#,
        ["cause"]
    };

    catch_filters_by_instanceof => {
        r#"Promise.reject(new TypeError("t")).catch(e=>e instanceof TypeError?"typed":"other").then(v=>console.log(v));"#,
        ["typed"]
    };

    catch_filters_non_matching_rethrows => {
        r#"Promise.reject(new TypeError("t")).catch(e=>{if(e instanceof RangeError)return "r";throw e;}).catch(e=>console.log(e.name));"#,
        ["TypeError"]
    };

    then_throw_in_arrow_vs_function_same => {
        r#"const o=[];Promise.resolve(0).then(function(){throw "fn";}).catch(e=>o.push(e));Promise.resolve(0).then(()=>{throw "ar";}).catch(e=>o.push(e)).then(()=>console.log(o.join(",")));"#,
        ["fn", "ar"]
    };

    promise_reject_after_resolve_in_executor_ignored => {
        r#"new Promise((res,rej)=>{res(1);rej(2);}).then(v=>console.log(v));"#,
        ["1"]
    };

    thenable_with_finally_like_method_ignored => {
        r#"Promise.resolve({then(res){res(4);},finally(){throw "no";}}).then(v=>console.log(v));"#,
        ["4"]
    };

    rejection_propagates_through_empty_then => {
        r#"Promise.reject("e").then().catch(e=>console.log(e));"#,
        ["e"]
    };

    catch_return_thenable_assimilated => {
        r#"Promise.reject("x").catch(()=>({then(res){res("assim");}})).then(v=>console.log(v));"#,
        ["assim"]
    };

    then_throw_after_promise_all_member_reject => {
        r#"Promise.all([Promise.reject("m")]).catch(e=>e).then(v=>{throw "wrap:"+v;}).catch(e=>console.log(e));"#,
        ["wrap:m"]
    };

    sequential_reject_catch_pairs => {
        r#"const o=[];Promise.reject(1).catch(e=>o.push("a:"+e));Promise.reject(2).catch(e=>o.push("b:"+e));Promise.resolve().then(()=>console.log(o.join(",")));"#,
        ["a:1,b:2"]
    };

    then_reject_handler_on_resolved_not_called => {
        r#"const o=[];Promise.resolve(3).then(v=>o.push("f:"+v),e=>o.push("r:"+e)).then(()=>console.log(o.join(",")));"#,
        ["f:3"]
    };

    promise_resolve_thenable_rejecting_promise_in_res => {
        r#"Promise.resolve({then(res){res(Promise.reject("pr"));}}).catch(e=>console.log(e));"#,
        ["pr"]
    };

    throw_after_catch_in_same_chain_branch => {
        r#"Promise.reject("a").catch(()=>"b").then(v=>{if(v==="b")throw "c";}).catch(e=>console.log(e));"#,
        ["c"]
    };

    then_multiple_handlers_only_first_runs => {
        r#"const o=[];const p=Promise.resolve(1);p.then(v=>o.push("h1:"+v));p.then(v=>o.push("h2:"+v));p.then(()=>console.log(o.join(",")));"#,
        ["h1:1,h2:1"]
    };

    rejection_skips_intermediate_thens_until_catch => {
        r#"const o=[];Promise.reject(0).then(()=>o.push("t1")).then(()=>o.push("t2")).catch(()=>o.push("c")).then(()=>console.log(o.join(",")));"#,
        ["c"]
    };

    catch_at_end_of_long_chain_recovers_once => {
        r#"Promise.resolve(1).then(x=>x).then(x=>x).then(()=>Promise.reject("deep")).then(x=>x).catch(e=>console.log(e));"#,
        ["deep"]
    };

    then_throw_eval_error => {
        r#"Promise.resolve(0).then(()=>{throw new EvalError("eval");}).catch(e=>console.log(e.name));"#,
        ["EvalError"]
    };

    then_throw_uri_error => {
        r#"Promise.resolve(0).then(()=>{throw new URIError("uri");}).catch(e=>console.log(e.name));"#,
        ["URIError"]
    };

    promise_reject_aggregate_error_reason => {
        r#"Promise.reject(new AggregateError([1,2],"agg")).catch(e=>console.log(e.errors.length));"#,
        ["2"]
    };

    catch_returns_object_reason => {
        r#"Promise.reject("x").catch(()=>({ok:true})).then(v=>console.log(v.ok));"#,
        ["true"]
    };

    thenable_sync_throw_before_executor_returns => {
        r#"Promise.resolve({then(){throw "syncThen";}}).catch(e=>console.log(e));"#,
        ["syncThen"]
    };

    rejection_from_returned_async_function_promise => {
        r#"Promise.resolve(0).then(()=>(async()=>{throw "af";})()).catch(e=>console.log(e));"#,
        ["af"]
    };

    then_catch_finally_catch_on_finally_throw => {
        r#"Promise.reject("a").catch(e=>e).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    promise_resolve_thenable_nested_reject => {
        r#"Promise.resolve({then(res){res({then(_,r){r("inner");}});}}).catch(e=>console.log(e));"#,
        ["inner"]
    };

    catch_handler_not_called_on_previous_recovery => {
        r#"Promise.reject("a").catch(()=>"ok").catch(e=>console.log("never:"+e)).then(v=>console.log(v));"#,
        ["ok"]
    };

    then_on_rejected_recovers_without_catch_method => {
        r#"Promise.reject(5).then(null,e=>e*2).then(v=>console.log(v));"#,
        ["10"]
    };

    throw_in_first_of_chained_thens_only => {
        r#"Promise.resolve(0).then(()=>{throw "only";}).then(()=>console.log("skip")).catch(e=>console.log(e));"#,
        ["only"]
    };

    promise_reject_with_error_subclass_message => {
        r#"class E extends Error{}Promise.reject(new E("sub")).catch(e=>console.log(e.message));"#,
        ["sub"]
    };

    then_return_value_after_catch_in_parent_branch => {
        r#"const o=[];Promise.reject("p").catch(e=>o.push(e));Promise.resolve("s").then(v=>o.push(v)).then(()=>console.log(o.join(",")));"#,
        ["p,s"]
    };

    rejection_bubbles_past_then_without_handler => {
        r#"Promise.reject("up").then().then().catch(e=>console.log(e));"#,
        ["up"]
    };

    catch_then_catch_double_recovery => {
        r#"Promise.reject("a").catch(e=>{throw "b:"+e;}).catch(e=>console.log(e));"#,
        ["b:a"]
    };

    thenable_promise_resolve_with_throw_in_fulfill_callback => {
        r#"new Promise(res=>res(1)).then(()=>{throw "cb";}).catch(e=>console.log(e));"#,
        ["cb"]
    };

    promise_resolve_immediate_then_throw => {
        r#"Promise.resolve().then(()=>{throw 0;}).catch(e=>console.log(e));"#,
        ["0"]
    };

    rejection_reason_type_number_zero => {
        r#"Promise.reject(0).catch(e=>console.log(e===0));"#,
        ["true"]
    };

    rejection_reason_type_empty_string => {
        r#"Promise.reject("").catch(e=>console.log(e.length));"#,
        ["0"]
    };

    then_throw_after_resolve_with_object => {
        r#"Promise.resolve({x:1}).then(o=>{throw o.x;}).catch(e=>console.log(e));"#,
        ["1"]
    };

    catch_identity_passthrough_rethrow => {
        r#"Promise.reject("id").catch(e=>Promise.reject(e)).catch(e=>console.log(e));"#,
        ["id"]
    };

    then_handler_throw_with_message_template => {
        r#"Promise.resolve("world").then(w=>{throw new Error(`hello ${w}`);}).catch(e=>console.log(e.message));"#,
        ["hello world"]
    };

    promise_allsettled_does_not_replace_rejection_propagation => {
        r#"Promise.allSettled([Promise.reject("s")]).then(r=>console.log(r[0].status));"#,
        ["rejected"]
    };

    then_catch_preserves_custom_error_name => {
        r#"class NamedErr extends Error{constructor(m,n){super(m);this.name=n;}}Promise.resolve(0).then(()=>{throw new NamedErr("d","AbortError");}).catch(e=>console.log(e.name));"#,
        ["AbortError"]
    };

    then_throw_symbol_primitive_caught => {
        r#"const s = Symbol("err_sym"); Promise.resolve().then(() => { throw s; }).catch(e => console.log(e === s));"#,
        ["true"]
    };
}
