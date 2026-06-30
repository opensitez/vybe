//! Promise job queue and microtask ordering — then/catch/finally scheduling.

crate::js_cases! {
    promise_then_runs_as_microtask_after_sync => {
        r#"const o=[]; Promise.resolve().then(()=>o.push("m")); o.push("s"); console.log(o.join(","));"#,
        ["s", "m"]
    };

    queue_microtask_before_promise_then => {
        r#"const o=[]; queueMicrotask(()=>o.push("q")); Promise.resolve().then(()=>o.push("p")); o.push("s"); console.log(o.join(","));"#,
        ["s", "q,p"]
    };

    nested_promise_then_order_fifo => {
        r#"const o=[]; Promise.resolve().then(()=>o.push("a")).then(()=>o.push("b")); o.push("s"); console.log(o.join(","));"#,
        ["s", "a,b"]
    };

    catch_runs_after_rejected_then => {
        r#"const o=[]; Promise.reject("e").catch(()=>o.push("c")).then(()=>o.push("t")); o.push("s"); console.log(o.join(","));"#,
        ["s", "c,t"]
    };

    finally_runs_after_fulfilled_then => {
        r#"const o=[]; Promise.resolve(1).then(()=>o.push("t")).finally(()=>o.push("f")); o.push("s"); console.log(o.join(","));"#,
        ["s", "t,f"]
    };

    finally_runs_after_rejection_catch => {
        r#"const o=[]; Promise.reject(1).catch(()=>o.push("c")).finally(()=>o.push("f")); o.push("s"); console.log(o.join(","));"#,
        ["s", "c,f"]
    };

    then_throw_schedules_rejection_handler => {
        r#"const o=[]; Promise.resolve().then(()=>{throw "x";}).catch(()=>o.push("c")); o.push("s"); console.log(o.join(","));"#,
        ["s", "c"]
    };

    promise_resolve_thenable_calls_then_async => {
        r#"const o=[]; Promise.resolve({then(cb){queueMicrotask(()=>cb(1));}}).then(v=>o.push(v)); o.push("s"); console.log(o.join(","));"#,
        ["s", "1"]
    };

    promise_all_waits_for_all_fulfilled => {
        r#"Promise.all([Promise.resolve(1),Promise.resolve(2)]).then(a=>console.log(a.join(",")));"#,
        ["1,2"]
    };

    promise_all_rejects_on_first_rejection => {
        r#"Promise.all([Promise.resolve(1),Promise.reject("f")]).catch(e=>console.log(e));"#,
        ["f"]
    };

    promise_all_settled_never_rejects => {
        r#"Promise.allSettled([Promise.resolve(1),Promise.reject("x")]).then(r=>console.log(r[1].status));"#,
        ["rejected"]
    };

    promise_race_first_settled_wins => {
        r#"Promise.race([new Promise(r=>setTimeout(()=>r("slow"),10)),Promise.resolve("fast")]).then(v=>console.log(v));"#,
        ["fast"]
    };

    promise_any_first_fulfillment_wins => {
        r#"Promise.any([Promise.reject("a"),Promise.resolve("ok")]).then(v=>console.log(v));"#,
        ["ok"]
    };

    async_await_schedules_continuation_microtask => {
        r#"const o=[]; (async()=>{o.push("a"); await Promise.resolve(); o.push("b");})(); o.push("s"); console.log(o.join(","));"#,
        ["s", "a,b"]
    };

    multiple_queue_microtask_fifo => {
        r#"const o=[]; queueMicrotask(()=>o.push("1")); queueMicrotask(()=>o.push("2")); o.push("s"); console.log(o.join(","));"#,
        ["s", "1,2"]
    };

    promise_chain_interleaved_with_sync => {
        r#"const o=[]; Promise.resolve().then(()=>o.push(1)).then(()=>o.push(2)); Promise.resolve().then(()=>o.push(3)); o.push(0); console.log(o.join(","));"#,
        ["0", "1,3,2"]
    };

    catch_return_value_becomes_fulfillment => {
        r#"Promise.reject("e").catch(()=>"ok").then(v=>console.log(v));"#,
        ["ok"]
    };

    finally_return_does_not_change_resolution => {
        r#"Promise.resolve("keep").finally(()=>"drop").then(v=>console.log(v));"#,
        ["keep"]
    };

    promise_finally_on_rejected_without_catch => {
        r#"Promise.resolve({then(){throw new Error("t");}}).catch(e=>console.log(e.message));"#,
        ["t"]
    };

    promise_resolve_identity_on_promise => {
        r#"const p=Promise.resolve(1); Promise.resolve(p).then(v=>console.log(v===1));"#,
        ["true"]
    };

    promise_reject_creates_rejected_promise => {
        r#"Promise.reject("no").catch(e=>console.log(e));"#,
        ["no"]
    };

    async_function_return_value_wrapped => {
        r#"(async()=>42)().then(v=>console.log(v));"#,
        ["42"]
    };

    async_throw_returns_rejected_promise => {
        r#"(async()=>{throw "ae";})().catch(e=>console.log(e));"#,
        ["ae"]
    };

    await_rejection_caught_in_async => {
        r#"(async()=>{try{await Promise.reject("ar");}catch(e){console.log(e);}})();"#,
        ["ar"]
    };

    promise_then_second_argument_handles_reject => {
        r#"Promise.reject("r").then(null,e=>console.log("h:"+e));"#,
        ["h:r"]
    };

    promise_resolve_in_then_creates_new_promise => {
        r#"Promise.resolve(1).then(()=>Promise.resolve(2)).then(v=>console.log(v));"#,
        ["2"]
    };

    microtask_from_promise_constructor_executor => {
        r#"const o=[]; new Promise(r=>{o.push("ex"); r();}).then(()=>o.push("th")); o.push("s"); console.log(o.join(","));"#,
        ["ex,s", "th"]
    };

    promise_all_empty_array_fulfills => {
        r#"Promise.all([]).then(a=>console.log(a.length));"#,
        ["0"]
    };

    promise_race_empty_never_settles_in_spec => {
        r#"Promise.race([]).then(()=>console.log("no"),()=>console.log("no")); console.log("sync");"#,
        ["sync"]
    };

    promise_with_resolvers_manual_resolve => {
        r#"const {promise,resolve}=Promise.withResolvers(); resolve(9); promise.then(v=>console.log(v));"#,
        ["9"]
    };

    promise_with_resolvers_manual_reject => {
        r#"const {promise,reject}=Promise.withResolvers(); reject("x"); promise.catch(e=>console.log(e));"#,
        ["x"]
    };

    then_catch_finally_order_on_rejection => {
        r#"const o=[]; Promise.reject(1).catch(()=>o.push("c")).then(()=>o.push("t")).finally(()=>o.push("f")); o.push("s"); console.log(o.join(","));"#,
        ["s", "c,t,f"]
    };

    nested_async_await_order => {
        r#"const o=[]; (async()=>{await Promise.resolve(); o.push("a"); await Promise.resolve(); o.push("b");})(); o.push("s"); console.log(o.join(","));"#,
        ["s", "a,b"]
    };

    promise_then_on_already_resolved_runs_microtask => {
        r#"const p=Promise.resolve(1); const o=[]; p.then(()=>o.push("t")); o.push("s"); console.log(o.join(","));"#,
        ["s", "t"]
    };

    catch_on_already_rejected_runs_microtask => {
        r#"const p=Promise.reject(1); const o=[]; p.catch(()=>o.push("c")); o.push("s"); console.log(o.join(","));"#,
        ["s", "c"]
    };

    promise_finally_on_rejected_runs_before_catch => {
        r#"Promise.reject("e").finally(()=>console.log("f")).catch(()=>console.log("c"));"#,
        ["f", "c"]
    };

    async_return_await_promise_value => {
        r#"(async()=>{return await Promise.resolve("v");})().then(x=>console.log(x));"#,
        ["v"]
    };

    queue_microtask_throw_caught_by_global_handler_pattern => {
        r#"const o=[]; queueMicrotask(()=>{try{throw "m";}catch(e){o.push(e);}}); o.push("s"); console.log(o.join(","));"#,
        ["s", "m"]
    };

    promise_all_settled_fulfilled_value_shape => {
        r#"Promise.allSettled([Promise.resolve(5)]).then(r=>console.log(r[0].value));"#,
        ["5"]
    };

    promise_any_all_reject_gives_aggregate => {
        r#"Promise.any([Promise.reject("a"),Promise.reject("b")]).catch(e=>console.log(e instanceof AggregateError));"#,
        ["true"]
    };
}
