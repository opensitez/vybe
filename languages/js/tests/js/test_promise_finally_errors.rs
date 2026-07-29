//! Promise .finally error semantics — throw in finally, finally after catch,
//! return in finally overriding rejection.

crate::js_cases! {
    finally_throw_on_resolved_promise => {
        r#"Promise.resolve(1).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_throw_on_rejected_promise => {
        r#"Promise.reject("r").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_throw_typeerror_replaces_resolve => {
        r#"Promise.resolve(0).finally(()=>{throw new TypeError("ft");}).catch(e=>console.log(e.name));"#,
        ["TypeError"]
    };

    finally_throw_after_catch_recovery => {
        r#"Promise.reject("a").catch(()=>"ok").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_throw_after_catch_reraise => {
        r#"Promise.reject("a").catch(e=>{throw "b:"+e;}).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_overrides_resolved_value => {
        r#"Promise.resolve("orig").finally(()=>"fin").then(v=>console.log(v));"#,
        ["orig"]
    };

    finally_return_overrides_rejection => {
        r#"Promise.reject("rej").finally(()=>"fin").catch(e=>console.log(e));"#,
        ["rej"]
    };

    finally_return_promise_overrides_reject => {
        r#"Promise.reject("r").finally(()=>Promise.resolve("fp")).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_return_rejected_promise_overrides_resolve => {
        r#"Promise.resolve(1).finally(()=>Promise.reject("fr")).catch(e=>console.log(e));"#,
        ["fr"]
    };

    finally_after_catch_then_finally_return => {
        r#"Promise.reject("e").catch(()=>"c").finally(()=>"f").then(v=>console.log(v));"#,
        ["c"]
    };

    finally_after_catch_then_finally_throw => {
        r#"Promise.reject("e").catch(()=>"c").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_runs_after_catch_without_rethrow => {
        r#"const o=[];Promise.reject("x").catch(e=>o.push("c:"+e)).finally(()=>o.push("f")).then(()=>console.log(o.join(",")));"#,
        ["c:x,f"]
    };

    finally_runs_on_resolve_without_throw => {
        r#"const o=[];Promise.resolve(1).finally(()=>o.push("f")).then(v=>o.push("v:"+v)).then(()=>console.log(o.join(",")));"#,
        ["f,v:1"]
    };

    finally_runs_on_reject_without_throw => {
        r#"const o=[];Promise.reject("r").finally(()=>o.push("f")).catch(e=>o.push("c:"+e)).then(()=>console.log(o.join(",")));"#,
        ["f,c:r"]
    };

    finally_does_not_swallow_prior_rejection => {
        r#"Promise.reject("keep").finally(()=>{}).catch(e=>console.log(e));"#,
        ["keep"]
    };

    finally_throw_supersedes_prior_resolution => {
        r#"Promise.resolve("ok").finally(()=>{throw "new";}).catch(e=>console.log(e));"#,
        ["new"]
    };

    finally_return_number_overrides => {
        r#"Promise.resolve("s").finally(()=>42).then(v=>console.log(v));"#,
        ["s"]
    };

    finally_return_object_overrides_reject => {
        r#"Promise.reject("x").finally(()=>({ok:1})).catch(e=>console.log(e));"#,
        ["x"]
    };

    chained_finally_first_throws => {
        r#"Promise.resolve(0).finally(()=>{throw "f1";}).finally(()=>"f2").catch(e=>console.log(e));"#,
        ["f1"]
    };

    chained_finally_second_throws => {
        r#"Promise.resolve(0).finally(()=>{}).finally(()=>{throw "f2";}).catch(e=>console.log(e));"#,
        ["f2"]
    };

    chained_finally_both_return_last_wins => {
        r#"Promise.resolve(0).finally(()=>"a").finally(()=>"b").then(v=>console.log(v));"#,
        ["0"]
    };

    finally_after_then_throw => {
        r#"Promise.resolve(1).then(()=>{throw "t";}).finally(()=>{}).catch(e=>console.log(e));"#,
        ["t"]
    };

    finally_before_catch_not_possible_order => {
        r#"Promise.reject("r").finally(()=>"f").catch(e=>"c:"+e).then(v=>console.log(v));"#,
        ["c:r"]
    };

    finally_return_undefined_overrides => {
        r#"Promise.resolve(5).finally(()=>{}).then(v=>console.log(v===undefined));"#,
        ["false"]
    };

    finally_throw_null_reason => {
        r#"Promise.resolve(0).finally(()=>{throw null;}).catch(e=>console.log(e===null));"#,
        ["true"]
    };

    finally_throw_undefined_reason => {
        r#"Promise.resolve(0).finally(()=>{throw undefined;}).catch(e=>console.log(e===undefined));"#,
        ["true"]
    };

    finally_on_promise_from_catch_recovery => {
        r#"Promise.reject("a").catch(()=>Promise.resolve("b")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_after_catch_on_reject => {
        r#"Promise.reject("err").catch(()=>"rec").finally(()=>"fin").then(v=>console.log(v));"#,
        ["rec"]
    };

    finally_async_callback_throw => {
        r#"Promise.resolve(1).finally(async()=>{throw "af";}).catch(e=>console.log(e));"#,
        ["af"]
    };

    finally_async_callback_return => {
        r#"Promise.reject("r").finally(async()=>"af").catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_return_promise_that_rejects => {
        r#"Promise.resolve(1).finally(()=>Promise.reject("inner")).catch(e=>console.log(e));"#,
        ["inner"]
    };

    finally_return_promise_that_resolves => {
        r#"Promise.reject("r").finally(()=>Promise.resolve("inner")).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_on_already_caught_promise => {
        r#"Promise.reject("x").catch(()=>{}).finally(()=>"f").then(v=>console.log(v));"#,
        ["undefined"]
    };

    finally_throw_error_with_message => {
        r#"Promise.resolve(0).finally(()=>{throw new Error("fm");}).catch(e=>console.log(e.message));"#,
        ["fm"]
    };

    finally_after_long_chain_reject => {
        r#"Promise.resolve(1).then(x=>x).then(()=>Promise.reject("deep")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_string_on_number_resolve => {
        r#"Promise.resolve(99).finally(()=>"str").then(v=>console.log(typeof v));"#,
        ["number"]
    };

    finally_throw_in_nested_function => {
        r#"Promise.resolve(0).finally(function(){throw "nf";}).catch(e=>console.log(e));"#,
        ["nf"]
    };

    finally_return_from_arrow_overrides => {
        r#"Promise.reject("r").finally(()=>"arr").catch(e=>console.log(e));"#,
        ["r"]
    };

    multiple_finally_on_same_branch => {
        r#"const o=[];Promise.resolve(1).finally(()=>o.push("f1")).finally(()=>o.push("f2")).then(()=>console.log(o.join(",")));"#,
        ["f1,f2"]
    };

    finally_after_catch_returns_promise => {
        r#"Promise.reject("a").catch(()=>Promise.resolve("c")).finally(()=>"f").then(v=>console.log(v));"#,
        ["c"]
    };

    finally_throw_range_error => {
        r#"Promise.resolve(0).finally(()=>{throw new RangeError("fr");}).catch(e=>console.log(e.name));"#,
        ["RangeError"]
    };

    finally_return_boolean_on_reject => {
        r#"Promise.reject("x").finally(()=>false).catch(e=>console.log(e));"#,
        ["x"]
    };

    finally_on_promise_all_rejection => {
        r#"Promise.all([Promise.reject("a")]).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_on_promise_race_rejection => {
        r#"Promise.race([Promise.reject("a")]).finally(()=>"rf").catch(e=>console.log(e));"#,
        ["a"]
    };

    finally_preserves_state_when_no_return_no_throw => {
        r#"Promise.resolve("keep").finally(()=>{}).then(v=>console.log(v));"#,
        ["keep"]
    };

    finally_reject_state_preserved_when_no_return => {
        r#"Promise.reject("lost").finally(()=>{}).catch(e=>console.log(e));"#,
        ["lost"]
    };

    finally_throw_after_finally_no_throw => {
        r#"Promise.resolve(1).finally(()=>{}).finally(()=>{throw "s";}).catch(e=>console.log(e));"#,
        ["s"]
    };

    finally_return_overrides_catch_scalar => {
        r#"Promise.reject("a").catch(()=>"c").finally(()=>"f").then(v=>console.log(v));"#,
        ["c"]
    };

    finally_then_finally_throw => {
        r#"Promise.resolve(1).then(x=>x).finally(()=>{}).finally(()=>{throw "t";}).catch(e=>console.log(e));"#,
        ["t"]
    };

    finally_with_side_effect_no_override => {
        r#"const o=[];Promise.resolve(7).finally(()=>o.push("s")).then(v=>console.log(v+","+o.join("")));"#,
        ["7,s"]
    };

    finally_throw_object_reason => {
        r#"Promise.resolve(0).finally(()=>{throw {code:1};}).catch(e=>console.log(e.code));"#,
        ["1"]
    };

    finally_return_thenable_assimilated => {
        r#"Promise.reject("r").finally(()=>({then(res){res("assim");}})).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_return_thenable_that_rejects => {
        r#"Promise.resolve(1).finally(()=>({then(_,r){r("tr");}})).catch(e=>console.log(e));"#,
        ["tr"]
    };

    finally_after_then_catch_recovery => {
        r#"Promise.resolve(0).then(()=>{throw "t";}).catch(()=>"ok").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_after_then_catch_finally_return => {
        r#"Promise.resolve(0).then(()=>{throw "t";}).catch(()=>"ok").finally(()=>"f").then(v=>console.log(v));"#,
        ["ok"]
    };

    finally_on_constructor_reject => {
        r#"new Promise((_,r)=>r("ctor")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_on_constructor_resolve_return => {
        r#"new Promise(res=>res("ctor")).finally(()=>"f").then(v=>console.log(v));"#,
        ["ctor"]
    };

    finally_throw_supersedes_catch_return => {
        r#"Promise.reject("a").catch(()=>"c").finally(()=>{throw "f";}).then(()=>console.log("skip")).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_supersedes_catch_return => {
        r#"Promise.reject("a").catch(()=>"c").finally(()=>"f").then(v=>console.log(v));"#,
        ["c"]
    };

    finally_second_in_chain_return_wins => {
        r#"Promise.resolve(0).finally(()=>"one").finally(()=>"two").then(v=>console.log(v));"#,
        ["0"]
    };

    finally_with_throw_after_catch_on_nested_reject => {
        r#"Promise.reject("n").catch(()=>Promise.reject("c")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_after_nested_catch_reject => {
        r#"Promise.reject("n").catch(()=>Promise.reject("c")).finally(()=>"f").catch(e=>console.log(e));"#,
        ["c"]
    };

    finally_on_identity_resolved_promise => {
        r#"const p=Promise.resolve(3);p.finally(()=>"x").then(v=>console.log(v));"#,
        ["3"]
    };

    finally_throw_reference_error => {
        r#"Promise.resolve(0).finally(()=>{throw new ReferenceError("ref");}).catch(e=>console.log(e.name));"#,
        ["ReferenceError"]
    };

    finally_return_bigint => {
        r#"Promise.reject("r").finally(()=>9n).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_throw_after_resolve_with_value => {
        r#"Promise.resolve({v:1}).finally(v=>{throw v===undefined?"noarg":"got";}).catch(e=>console.log(e));"#,
        ["noarg"]
    };

    finally_no_op_between_catch_and_then => {
        r#"Promise.reject("e").catch(()=>"c").finally(()=>{}).then(v=>console.log(v));"#,
        ["c"]
    };

    finally_return_empty_string => {
        r#"Promise.resolve("x").finally(()=>"").then(v=>console.log(v.length));"#,
        ["1"]
    };

    finally_throw_in_conditional => {
        r#"Promise.resolve(true).finally(ok=>{throw ok?"truthy":"noarg";}).catch(e=>console.log(e));"#,
        ["noarg"]
    };

    finally_return_in_conditional => {
        r#"Promise.reject("r").finally(()=>{return "yes";}).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_on_delayed_reject => {
        r#"new Promise((_,r)=>setTimeout(()=>r("late"),0)).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_on_delayed_resolve_return => {
        r#"new Promise(res=>setTimeout(()=>res("late"),0)).finally(()=>"f").then(v=>console.log(v));"#,
        ["late"]
    };

    finally_after_multiple_catches => {
        r#"Promise.reject("a").catch(e=>{throw "b:"+e;}).catch(()=>"c").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_after_multiple_catches => {
        r#"Promise.reject("a").catch(e=>{throw "b:"+e;}).catch(()=>"c").finally(()=>"f").then(v=>console.log(v));"#,
        ["c"]
    };

    finally_throw_custom_error_subclass => {
        r#"class E extends Error{}Promise.resolve(0).finally(()=>{throw new E("ce");}).catch(e=>console.log(e instanceof E));"#,
        ["true"]
    };

    finally_return_promise_resolving_to_object => {
        r#"Promise.reject("r").finally(()=>Promise.resolve({k:1})).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_on_chain_with_interleaved_then => {
        r#"Promise.resolve(1).then(x=>x+1).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_on_chain_with_interleaved_then_return => {
        r#"Promise.resolve(1).then(x=>x+1).finally(()=>"f").then(v=>console.log(v));"#,
        ["2"]
    };

    finally_after_catch_throws => {
        r#"Promise.reject("a").catch(()=>{throw "c";}).finally(()=>"f").catch(e=>console.log(e));"#,
        ["c"]
    };

    finally_runs_even_when_catch_throws_before_finally => {
        r#"const o=[];Promise.reject("a").catch(()=>{throw "c";}).finally(()=>o.push("f")).catch(e=>o.push("e:"+e)).then(()=>console.log(o.join(",")));"#,
        ["f,e:c"]
    };

    finally_return_zero_overrides_truthy_resolve => {
        r#"Promise.resolve(1).finally(()=>0).then(v=>console.log(v));"#,
        ["1"]
    };

    finally_throw_false_reason => {
        r#"Promise.resolve(0).finally(()=>{throw false;}).catch(e=>console.log(String(e)));"#,
        ["false"]
    };

    finally_on_promise_resolve_thenable => {
        r#"Promise.resolve({then(res){res(1);}}).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_on_promise_resolve_thenable_return => {
        r#"Promise.resolve({then(res){res(1);}}).finally(()=>"f").then(v=>console.log(v));"#,
        ["1"]
    };

    finally_after_then_return_rejected => {
        r#"Promise.resolve(0).then(()=>Promise.reject("tr")).finally(()=>"f").catch(e=>console.log(e));"#,
        ["tr"]
    };

    finally_throw_after_then_return_rejected => {
        r#"Promise.resolve(0).then(()=>Promise.reject("tr")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_with_explicit_undefined_return => {
        r#"Promise.resolve(5).finally(()=>undefined).then(v=>console.log(v===undefined));"#,
        ["false"]
    };

    finally_passes_through_when_callback_missing_behavior => {
        r#"Promise.resolve("p").finally(()=>null).then(v=>console.log(v===null));"#,
        ["false"]
    };

    finally_on_reject_catch_finally_order => {
        r#"const o=[];Promise.reject("r").catch(e=>o.push("c")).finally(()=>o.push("f")).then(()=>console.log(o.join(",")));"#,
        ["c,f"]
    };

    finally_throw_replaces_finally_return_in_same => {
        r#"Promise.resolve(0).finally(()=>{throw "t";}).catch(e=>console.log(e));"#,
        ["t"]
    };

    finally_return_array => {
        r#"Promise.reject("r").finally(()=>[1,2]).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_throw_in_loop_inside_callback => {
        r#"Promise.resolve(0).finally(()=>{for(let i=0;i<2;i++){if(i===1)throw "loop";}}).catch(e=>console.log(e));"#,
        ["loop"]
    };

    finally_return_from_nested_try => {
        r#"Promise.reject("r").finally(()=>{try{return "inner";}catch{return "bad";}}).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_throw_from_nested_try => {
        r#"Promise.resolve(0).finally(()=>{try{throw "inner";}catch(e){throw "wrap:"+e;}}).catch(e=>console.log(e));"#,
        ["wrap:inner"]
    };

    finally_on_allsettled_result => {
        r#"Promise.allSettled([Promise.reject("a")]).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_on_allsettled => {
        r#"Promise.allSettled([Promise.reject("a")]).finally(()=>"fs").then(v=>console.log(v[0].status));"#,
        ["rejected"]
    };

    finally_after_sync_throw_in_then => {
        r#"Promise.resolve(0).then(()=>{throw "s";}).finally(()=>"f").catch(e=>console.log(e));"#,
        ["s"]
    };

    finally_runs_after_sync_throw_before_catch => {
        r#"const o=[];Promise.resolve(0).then(()=>{throw "s";}).finally(()=>o.push("f")).catch(e=>o.push("c:"+e)).then(()=>console.log(o.join(",")));"#,
        ["f,c:s"]
    };

    finally_return_overrides_finally_side_effect_only => {
        r#"const o=[];Promise.resolve(1).finally(()=>{o.push("s");return "r";}).then(v=>console.log(v+o.join("")));"#,
        ["1s"]
    };

    finally_throw_overrides_pending_resolve_value => {
        r#"Promise.resolve("pending").finally(()=>{throw "override";}).catch(e=>console.log(e));"#,
        ["override"]
    };

    finally_return_symbol => {
        r#"const s=Symbol("f");Promise.reject("r").finally(()=>s).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_throw_symbol => {
        r#"const s=Symbol("e");Promise.resolve(0).finally(()=>{throw s;}).catch(e=>console.log(e===s));"#,
        ["true"]
    };

    finally_on_branching_catch_recovery => {
        r#"Promise.reject("x").catch(e=>e==="x"?"ok":"no").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_on_branching_catch => {
        r#"Promise.reject("x").catch(e=>e==="x"?"ok":"no").finally(()=>"f").then(v=>console.log(v));"#,
        ["ok"]
    };

    finally_after_promise_resolve_nested => {
        r#"Promise.resolve(Promise.resolve(1)).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_after_promise_resolve_nested => {
        r#"Promise.resolve(Promise.resolve(1)).finally(()=>"f").then(v=>console.log(v));"#,
        ["1"]
    };

    finally_on_reject_wrapped_error => {
        r#"Promise.reject(new Error("we")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_after_reject_wrapped_error => {
        r#"Promise.reject(new Error("we")).finally(()=>"f").catch(e=>console.log(e.message));"#,
        ["we"]
    };

    finally_three_deep_last_throw => {
        r#"Promise.resolve(0).finally(()=>{}).finally(()=>{}).finally(()=>{throw "3";}).catch(e=>console.log(e));"#,
        ["3"]
    };

    finally_three_deep_last_return => {
        r#"Promise.resolve(0).finally(()=>"a").finally(()=>"b").finally(()=>"c").then(v=>console.log(v));"#,
        ["0"]
    };

    finally_catch_on_finally_throw_separate_branch => {
        r#"Promise.resolve(1).finally(()=>{throw "a";}).catch(e=>"c:"+e).finally(()=>"f").then(v=>console.log(v));"#,
        ["c:a"]
    };

    finally_after_catch_on_finally_throw => {
        r#"Promise.resolve(1).finally(()=>{throw "a";}).catch(e=>e).finally(()=>"f").then(v=>console.log(v));"#,
        ["a"]
    };

    finally_return_promise_chain_in_callback => {
        r#"Promise.reject("r").finally(()=>Promise.resolve().then(()=>"chain")).catch(e=>console.log(e));"#,
        ["r"]
    };

    finally_throw_after_promise_in_callback => {
        r#"Promise.resolve(0).finally(()=>Promise.resolve().then(()=>{throw "chain";})).catch(e=>console.log(e));"#,
        ["chain"]
    };

    finally_on_mixed_resolve_reject_chain => {
        r#"Promise.resolve(1).then(()=>Promise.reject("m")).finally(()=>"f").catch(e=>console.log(e));"#,
        ["m"]
    };

    finally_throw_on_mixed_resolve_reject_chain => {
        r#"Promise.resolve(1).then(()=>Promise.reject("m")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_with_null_return_overrides => {
        r#"Promise.resolve("x").finally(()=>null).then(v=>console.log(v===null));"#,
        ["false"]
    };

    finally_callback_receives_no_args => {
        r#"Promise.resolve(99).finally((...a)=>a.length).then(v=>console.log(v));"#,
        ["99"]
    };

    finally_on_already_rejected_then_caught => {
        r#"Promise.reject("a").catch(()=>{}).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_on_already_rejected_then_caught => {
        r#"Promise.reject("a").catch(()=>{}).finally(()=>"f").then(v=>console.log(v));"#,
        ["undefined"]
    };

    finally_syntax_error_throw => {
        r#"Promise.resolve(0).finally(()=>{throw new SyntaxError("syn");}).catch(e=>console.log(e.name));"#,
        ["SyntaxError"]
    };

    finally_eval_error_throw => {
        r#"Promise.resolve(0).finally(()=>{throw new EvalError("ev");}).catch(e=>console.log(e.name));"#,
        ["EvalError"]
    };

    finally_uri_error_throw => {
        r#"Promise.resolve(0).finally(()=>{throw new URIError("ur");}).catch(e=>console.log(e.name));"#,
        ["URIError"]
    };

    finally_after_thenable_assimilation => {
        r#"Promise.resolve({then(res){res(5);}}).then(v=>v).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_after_thenable_assimilation => {
        r#"Promise.resolve({then(res){res(5);}}).then(v=>v).finally(()=>"f").then(v=>console.log(v));"#,
        ["5"]
    };

    finally_on_promise_from_async_then => {
        r#"Promise.resolve(0).then(async()=>1).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_on_promise_from_async_then => {
        r#"Promise.resolve(0).then(async()=>1).finally(()=>"f").then(v=>console.log(v));"#,
        ["1"]
    };

    finally_rethrows_after_catch_if_finally_throws => {
        r#"Promise.reject("a").catch(()=>"ok").finally(()=>{throw "f";}).catch(e=>console.log(e!=="ok"));"#,
        ["true"]
    };

    finally_does_not_run_catch_on_return_override => {
        r#"const o=[];Promise.reject("r").catch(e=>o.push("c")).finally(()=>"f").then(v=>o.push("t:"+v)).then(()=>console.log(o.join(",")));"#,
        ["c,t:1"]
    };

    finally_throw_aggregate_error => {
        r#"Promise.resolve(0).finally(()=>{throw new AggregateError([],"ag");}).catch(e=>console.log(e instanceof AggregateError));"#,
        ["true"]
    };

    finally_return_after_catch_identity => {
        r#"Promise.reject("id").catch(e=>e).finally(()=>"f").then(v=>console.log(v));"#,
        ["id"]
    };

    finally_on_parallel_branch_isolation => {
        r#"const o=[];Promise.resolve(1).finally(()=>o.push("a"));Promise.reject("b").finally(()=>o.push("b")).catch(()=>{});Promise.resolve().then(()=>console.log(o.sort().join(",")));"#,
        ["a,b"]
    };

    finally_throw_string_on_number_reject => {
        r#"Promise.reject(404).finally(()=>{throw "nf";}).catch(e=>console.log(e));"#,
        ["nf"]
    };

    finally_return_string_on_number_reject => {
        r#"Promise.reject(404).finally(()=>"nf").catch(e=>console.log(e));"#,
        ["404"]
    };

    finally_after_double_then_reject => {
        r#"Promise.resolve(0).then(()=>1).then(()=>Promise.reject("d")).finally(()=>"f").catch(e=>console.log(e));"#,
        ["d"]
    };

    finally_throw_after_double_then_reject => {
        r#"Promise.resolve(0).then(()=>1).then(()=>Promise.reject("d")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_with_void_operator_return => {
        r#"Promise.resolve(1).finally(()=>void "ignored").then(v=>console.log(v===undefined));"#,
        ["false"]
    };

    finally_return_negates_rejection_message => {
        r#"Promise.reject("msg").finally(()=>"recovered").catch(e=>console.log(e));"#,
        ["msg"]
    };

    finally_throw_in_finally_after_finally_return => {
        r#"Promise.resolve(0).finally(()=>"skip").finally(()=>{throw "win";}).catch(e=>console.log(e));"#,
        ["win"]
    };

    finally_return_in_finally_after_finally_throw_caught => {
        r#"Promise.resolve(0).finally(()=>{throw "a";}).catch(()=>"b").finally(()=>"c").then(v=>console.log(v));"#,
        ["b"]
    };

    finally_on_reject_empty_string_reason => {
        r#"Promise.reject("").finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_on_reject_empty_string => {
        r#"Promise.reject("").finally(()=>"f").catch(e=>console.log(e===""));"#,
        ["true"]
    };

    finally_promise_reject_in_executor_then_finally => {
        r#"new Promise((_,r)=>r("ex")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_promise_resolve_in_executor_then_finally_return => {
        r#"new Promise(res=>res("ex")).finally(()=>"f").then(v=>console.log(v));"#,
        ["ex"]
    };

    finally_after_catch_returning_rejected_promise => {
        r#"Promise.reject("a").catch(()=>Promise.reject("c")).finally(()=>"f").catch(e=>console.log(e));"#,
        ["c"]
    };

    finally_throw_after_catch_returning_rejected_promise => {
        r#"Promise.reject("a").catch(()=>Promise.reject("c")).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_on_then_with_onrejected => {
        r#"Promise.reject("r").then(null,e=>"h:"+e).finally(()=>{throw "f";}).catch(e=>console.log(e));"#,
        ["f"]
    };

    finally_return_on_then_with_onrejected => {
        r#"Promise.reject("r").then(null,e=>"h:"+e).finally(()=>"f").then(v=>console.log(v));"#,
        ["h:r"]
    };

    finally_completes_before_catch_on_throw => {
        r#"const o=[];Promise.resolve(1).finally(()=>o.push("f")).then(()=>{throw "t";}).catch(e=>o.push("c:"+e)).then(()=>console.log(o.join(",")));"#,
        ["f,c:t"]
    };

    finally_on_fulfilled_after_catch_in_sibling => {
        r#"const o=[];Promise.reject("a").catch(()=>{});Promise.resolve("b").finally(()=>o.push("f")).then(v=>o.push(v)).then(()=>console.log(o.join(",")));"#,
        ["f,b"]
    };

    finally_return_delayed_promise_resolving_value => {
        r#"Promise.resolve(10).finally(()=>new Promise(res=>res(20))).then(v=>console.log(v));"#,
        ["10"]
    };
}
