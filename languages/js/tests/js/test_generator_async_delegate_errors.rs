//! Generator and async-generator delegation — yield* throw/return/close propagation.

crate::js_cases! {
    yield_star_forwards_next_values => {
        r#"function* inner(){yield 1;yield 2;} function* outer(){yield* inner();} const g=outer(); console.log(g.next().value);console.log(g.next().value);"#,
        ["1", "2"]
    };

    yield_star_throw_propagates_to_outer_iterator => {
        r#"function* inner(){yield 1;throw new Error("inner");} function* outer(){try{yield* inner();}catch(e){yield e.message;}} const g=outer(); g.next(); console.log(g.next().value);"#,
        ["inner"]
    };

    yield_star_return_value_becomes_yield_star_result => {
        r#"function* inner(){yield 1;return "done";} function* outer(){const v=yield* inner(); yield v;} const g=outer(); g.next(); g.next(); console.log(g.next().value);"#,
        ["done"]
    };

    yield_star_on_string_iterates_code_units => {
        r#"function* g(){yield*"ab";} const a=[...g()]; console.log(a.join(""));"#,
        ["ab"]
    };

    yield_star_on_array_iterates_elements => {
        r#"function* g(){yield*[10,20];} const a=[...g()]; console.log(a.join(","));"#,
        ["10,20"]
    };

    generator_throw_resumes_at_throw_site => {
        r#"function* g(){try{yield 1;}catch(e){yield e;}} const gen=g(); gen.next(); console.log(gen.throw("x").value);"#,
        ["x"]
    };

    generator_return_closes_iterator => {
        r#"function* g(){yield 1;yield 2;} const gen=g(); gen.next(); console.log(gen.return(9).value);console.log(gen.next().done);"#,
        ["9", "true"]
    };

    generator_throw_after_completion_ignored => {
        r#"function* g(){yield 1;} const gen=g(); gen.next(); gen.next(); console.log(gen.throw("late").done);"#,
        ["true"]
    };

    async_generator_yield_star_async_inner => {
        r#"async function* inner(){yield 1;yield 2;} async function* outer(){yield* inner();} (async()=>{const a=[];for await(const v of outer())a.push(v);console.log(a.join(","));})();"#,
        ["1,2"]
    };

    async_generator_throw_inside_caught => {
        r#"async function* g(){try{yield 1;throw new Error("a");}catch(e){yield e.message;}} (async()=>{const a=[];for await(const v of g())a.push(v);console.log(a.join(","));})();"#,
        ["1,a"]
    };

    async_generator_return_stops_iteration => {
        r#"async function* g(){yield 1;return "end";yield 2;} (async()=>{const a=[];for await(const v of g())a.push(v);console.log(a.length);})();"#,
        ["1"]
    };

    yield_star_nested_three_levels => {
        r#"function* a(){yield 1;} function* b(){yield* a();} function* c(){yield* b();} console.log([...c()][0]);"#,
        ["1"]
    };

    yield_star_inner_return_outer_continues => {
        r#"function* inner(){return 5;} function* outer(){const r=yield* inner(); yield r+1;} console.log([...outer()][0]);"#,
        ["6"]
    };

    yield_star_with_manual_iterator_throw => {
        r#"const it={*[Symbol.iterator](){yield 1;throw new Error("it");}}; function* g(){try{yield* it;}catch(e){yield "caught";}} console.log([...g()].join(","));"#,
        ["1,caught"]
    };

    generator_send_value_to_yield_expression => {
        r#"function* g(){const x=yield 1; yield x;} const gen=g(); gen.next(); console.log(gen.next(9).value);"#,
        ["9"]
    };

    generator_yield_in_try_finally_preserves_flow => {
        r#"function* g(){try{yield 1;}finally{yield 2;}} console.log([...g()].join(","));"#,
        ["1,2"]
    };

    async_for_await_yield_star_async_gen => {
        r#"async function* nums(){yield 3;yield 4;} async function* wrap(){yield* nums();} (async()=>{let s=0;for await(const v of wrap())s+=v;console.log(s);})();"#,
        ["7"]
    };

    async_generator_delegate_throw_to_consumer => {
        r#"async function* g(){yield 1;throw new Error("stop");} (async()=>{const o=[];try{for await(const v of g())o.push(v);}catch(e){o.push(e.message);}console.log(o.join(","));})();"#,
        ["1,stop"]
    };

    yield_star_on_generator_object_directly => {
        r#"function* inner(){yield "x";} function* outer(){yield* inner();} console.log(outer().next().value);"#,
        ["x"]
    };

    generator_close_via_return_before_first_yield => {
        r#"function* g(){yield 1;} const gen=g(); console.log(gen.return(0).done);"#,
        ["true"]
    };

    yield_star_empty_generator_returns_undefined => {
        r#"function* empty(){} function* g(){const r=yield* empty(); yield r===undefined;} console.log([...g()][0]);"#,
        ["true"]
    };

    async_generator_await_inside_yield_star => {
        r#"async function* inner(){yield await Promise.resolve(2);} async function* outer(){yield* inner();} (async()=>{const v=await outer().next();console.log(v.value);})();"#,
        ["2"]
    };

    generator_delegate_to_custom_iterable_with_return => {
        r#"const iterable={[Symbol.iterator](){let n=0;return{next(){return n++?{value:undefined,done:true}:{value:7,done:false};},return(v){return{value:v,done:true};}};}}; function* g(){const r=yield* iterable; yield r;} console.log([...g()][0]);"#,
        ["undefined"]
    };

    yield_star_throw_from_nested_delegate => {
        r#"function* a(){throw "a";} function* b(){yield* a();} function* c(){try{yield* b();}catch(e){yield e;}} console.log([...c()][0]);"#,
        ["a"]
    };

    async_yield_star_promise_rejecting_async_iterable => {
        r#"async function* bad(){yield 1;await Promise.reject("nope");} async function* wrap(){try{yield* bad();}catch(e){yield "e:"+e;}} (async()=>{const a=[];for await(const v of wrap())a.push(v);console.log(a.join(","));})();"#,
        ["1,e:nope"]
    };

    generator_next_after_throw_still_closed => {
        r#"function* g(){throw new Error("e");} const gen=g(); try{gen.next();}catch{} console.log(gen.next().done);"#,
        ["true"]
    };

    yield_star_array_spread_in_generator_expression => {
        r#"const g=function*(){yield*[1,2,3];}; console.log([...g()].length);"#,
        ["3"]
    };

    async_generator_function_star_name_preserved => {
        r#"async function* stream(){} console.log(stream.name);"#,
        ["stream"]
    };

    generator_composed_with_map_on_output => {
        r#"function* nums(){yield 1;yield 2;} const mapped=[...nums()].map(x=>x*10); console.log(mapped.join(","));"#,
        ["10,20"]
    };

    yield_star_from_throw_in_inner_skips_outer_remaining => {
        r#"function* inner(){yield 1;throw new Error("x");yield 2;} function* outer(){try{yield* inner();}catch(e){yield "c";}} console.log([...outer()].join(","));"#,
        ["1,c"]
    };

    async_generator_throw_method_resumes => {
        r#"async function* g() { try { yield 1; } catch (e) { yield "caught:" + e; } } (async () => { const gen = g(); await gen.next(); const r = await gen.throw("err"); console.log(r.value); })();"#,
        ["caught:err"]
    };

    async_generator_return_method_resolves_value => {
        r#"(async () => { async function* g() { yield 1; yield 2; } const gen = g(); await gen.next(); const r = await gen.return("custom_ret"); console.log(r.value + "|" + r.done); })();"#,
        ["custom_ret|true"]
    };
}


