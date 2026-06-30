//! Async/await error recovery — await in loops with try/catch, parallel await
//! errors, async IIFE errors, for-await throw.

crate::js_cases! {
    async_for_loop_catch_each_rejection => {
        r#"async function main(){const o=[];for(let i=0;i<3;i++){try{await Promise.reject(i);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["0,1,2"]
    };

    async_for_loop_continue_after_catch => {
        r#"async function main(){const o=[];for(let i=0;i<3;i++){try{if(i===1)throw "skip";o.push(i);}catch{o.push("c");}}console.log(o.join(","));}main();"#,
        ["0,c,2"]
    };

    async_for_loop_break_on_caught_error => {
        r#"async function main(){const o=[];for(let i=0;i<5;i++){try{await Promise.reject(i);}catch(e){o.push(e);if(e===2)break;}}console.log(o.join(","));}main();"#,
        ["0,1,2"]
    };

    async_while_loop_catch_await_rejection => {
        r#"async function main(){let n=0;const o=[];while(n<3){try{await Promise.reject(n);}catch(e){o.push(e);n++;}}console.log(o.join(","));}main();"#,
        ["0,1,2"]
    };

    async_do_while_catch_await_throw => {
        r#"async function main(){const o=[];let i=0;do{try{await (async()=>{throw i;})();}catch(e){o.push(e);}i++;}while(i<2);console.log(o.join(","));}main();"#,
        ["0,1"]
    };

    async_for_of_array_await_with_catch => {
        r#"async function main(){const o=[];for(const x of [1,2,3]){try{await Promise.reject(x);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["1,2,3"]
    };

    async_for_in_object_await_rejection => {
        r#"async function main(){const o=[];const obj={a:1,b:2};for(const k in obj){try{await Promise.reject(k);}catch(e){o.push(e);}}console.log(o.sort().join(","));}main();"#,
        ["a,b"]
    };

    async_parallel_await_all_catch_first_error => {
        r#"async function main(){try{await Promise.all([Promise.resolve(1),Promise.reject("p"),Promise.resolve(3)]);}catch(e){console.log(e);}}main();"#,
        ["p"]
    };

    async_parallel_await_all_individual_try_catch => {
        r#"async function main(){const ps=[Promise.resolve(1),Promise.reject("b"),Promise.resolve(3)];const o=[];for(const p of ps){try{o.push(await p);}catch(e){o.push("e:"+e);}}console.log(o.join(","));}main();"#,
        ["1,e:b,3"]
    };

    async_iife_throws_before_await => {
        r#"async function main(){try{await (async()=>{throw "early";})();}catch(e){console.log(e);}}main();"#,
        ["early"]
    };

    async_iife_catches_own_throw => {
        r#"async function main(){const v=await (async()=>{try{throw "in";}catch(e){return "got:"+e;}})();console.log(v);}main();"#,
        ["got:in"]
    };

    async_iife_await_reject_uncaught => {
        r#"async function main(){try{await (async()=>{await Promise.reject("iife");})();}catch(e){console.log(e);}}main();"#,
        ["iife"]
    };

    async_arrow_iife_error_recovery => {
        r#"async function main(){const r=await (async()=>{try{await Promise.reject("x");return "no";}catch{return "yes";}})();console.log(r);}main();"#,
        ["yes"]
    };

    async_for_await_collect_until_throw => {
        r#"async function main(){const o=[];async function* gen(){yield 1;yield 2;throw "stop";yield 3;}try{for await(const v of gen())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,2,e:stop"]
    };

    async_for_await_empty_iterator_no_error => {
        r#"async function main(){const o=[];async function* empty(){}for await(const v of empty())o.push(v);console.log(o.length);}main();"#,
        ["0"]
    };

    async_for_await_async_iterable_reject => {
        r#"async function main(){const o=[];const it={async*[Symbol.asyncIterator](){yield 1;throw "iter";}};try{for await(const v of it)o.push(v);}catch(e){o.push("c:"+e);}console.log(o.join(","));}main();"#,
        ["1,c:iter"]
    };

    async_nested_function_await_catch => {
        r#"async function main(){async function inner(){try{await Promise.reject("in");}catch(e){return e;}}console.log(await inner());}main();"#,
        ["in"]
    };

    async_nested_try_catch_with_await => {
        r#"async function main(){try{try{await Promise.reject("a");}catch(e){throw "b:"+e;}}catch(e){console.log(e);}}main();"#,
        ["b:a"]
    };

    async_for_loop_labeled_break_in_catch => {
        r#"async function main(){const o=[];outer:for(let i=0;i<3;i++){try{await Promise.reject(i);}catch(e){o.push(e);break outer;}}console.log(o.join(","));}main();"#,
        ["0"]
    };

    async_for_loop_labeled_continue_in_catch => {
        r#"async function main(){const o=[];outer:for(let i=0;i<3;i++){try{await Promise.reject(i);}catch(e){o.push("c"+e);continue outer;o.push("skip");}}console.log(o.join(","));}main();"#,
        ["c0,c1,c2"]
    };

    async_switch_await_rejection_in_case => {
        r#"async function main(){const o=[];for(const k of [1,2]){switch(k){case 1:try{await Promise.reject("s");}catch(e){o.push(e);}break;case 2:o.push("ok");}}console.log(o.join(","));}main();"#,
        ["s,ok"]
    };

    async_ternary_await_reject_catch => {
        r#"async function main(){try{const v=await (true?Promise.reject("t"):Promise.resolve(1));console.log(v);}catch(e){console.log(e);}}main();"#,
        ["t"]
    };

    async_short_circuit_and_skips_await => {
        r#"async function main(){const v=false&&await Promise.reject("skip");console.log(v);}main();"#,
        ["false"]
    };

    async_short_circuit_or_awaits_reject => {
        r#"async function main(){try{const v=true||await Promise.reject("skip");console.log(v);}catch(e){console.log("err");}}main();"#,
        ["true"]
    };

    async_promise_all_settled_mixed_errors => {
        r#"async function main(){const r=await Promise.allSettled([Promise.resolve(1),Promise.reject("e")]);console.log(r[1].reason);}main();"#,
        ["e"]
    };

    async_sequential_await_recovery_between => {
        r#"async function main(){const o=[];for(const p of [Promise.reject("a"),Promise.resolve("b")]){try{o.push(await p);}catch(e){o.push("x");}}console.log(o.join(","));}main();"#,
        ["x,b"]
    };

    async_map_with_await_and_catch => {
        r#"async function main(){const xs=[1,2,3];const o=[];for(const x of xs){try{await Promise.reject(x);}catch(e){o.push(e*10);}}console.log(o.join(","));}main();"#,
        ["10,20,30"]
    };

    async_reduce_with_await_errors => {
        r#"async function main(){const o=[];let acc=0;for(const x of [1,2]){try{acc+=await Promise.resolve(x);}catch(e){o.push(e);}}console.log(acc);}main();"#,
        ["3"]
    };

    async_while_true_break_on_error => {
        r#"async function main(){const o=[];let i=0;while(true){try{await Promise.reject(i++);}catch(e){o.push(e);if(e===1)break;}}console.log(o.join(","));}main();"#,
        ["0,1"]
    };

    async_for_await_promise_array_reject => {
        r#"async function main(){const o=[];try{for await(const v of [Promise.resolve(1),Promise.reject("p")])o.push(await v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:p"]
    };

    async_class_method_await_catch => {
        r#"async function main(){class C{async run(){try{await Promise.reject("cm");}catch(e){return e;}}}console.log(await new C().run());}main();"#,
        ["cm"]
    };

    async_static_method_await_error => {
        r#"async function main(){class C{static async run(){try{await Promise.reject("st");}catch(e){return e;}}}console.log(await C.run());}main();"#,
        ["st"]
    };

    async_getter_await_rejection => {
        r#"async function main(){const o={async get val(){throw "getter";}};try{console.log(await o.val);}catch(e){console.log(e);}}main();"#,
        ["getter"]
    };

    async_try_finally_on_await_reject => {
        r#"async function main(){const o=[];try{await Promise.reject("r");}catch(e){o.push(e);}finally{o.push("f");}console.log(o.join(","));}main();"#,
        ["r,f"]
    };

    async_catch_return_skips_finally_override => {
        r#"async function main(){async function f(){try{await Promise.reject("x");}catch{return "c";}finally{return "f";}}console.log(await f());}main();"#,
        ["f"]
    };

    async_await_in_catch_block => {
        r#"async function main(){try{throw "sync";}catch(e){const v=await Promise.resolve("async:"+e);console.log(v);}}main();"#,
        ["async:sync"]
    };

    async_await_in_finally_after_catch => {
        r#"async function main(){const o=[];try{await Promise.reject("a");}catch(e){o.push(e);}finally{o.push(await Promise.resolve("f"));}console.log(o.join(","));}main();"#,
        ["a,f"]
    };

    async_parallel_two_awaits_both_caught => {
        r#"async function main(){const o=[];const a=(async()=>{try{await Promise.reject("1");}catch(e){o.push(e);}})();const b=(async()=>{try{await Promise.reject("2");}catch(e){o.push(e);}})();await Promise.all([a,b]);console.log(o.sort().join(","));}main();"#,
        ["1,2"]
    };

    async_race_await_rejection => {
        r#"async function main(){try{await Promise.race([Promise.reject("fast"),new Promise(()=>{})]);}catch(e){console.log(e);}}main();"#,
        ["fast"]
    };

    async_any_await_all_reject => {
        r#"async function main(){try{await Promise.any([Promise.reject("a"),Promise.reject("b")]);}catch(e){console.log(e instanceof AggregateError);}}main();"#,
        ["true"]
    };

    async_for_await_generator_return_early => {
        r#"async function main(){const o=[];async function* g(){yield 1;return 99;yield 2;}for await(const v of g())o.push(v);console.log(o.join(","));}main();"#,
        ["1"]
    };

    async_for_await_throw_on_second_yield => {
        r#"async function main(){const o=[];async function* g(){yield "a";throw "b";}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["a,e:b"]
    };

    async_iife_immediate_invoke_with_await => {
        r#"async function main(){const r=await(async()=>{try{await Promise.reject("q");}catch(e){return e+"!";}})();console.log(r);}main();"#,
        ["q!"]
    };

    async_nested_iife_error_bubbles => {
        r#"async function main(){try{await(async()=>await(async()=>{throw "deep";})())();}catch(e){console.log(e);}}main();"#,
        ["deep"]
    };

    async_for_loop_await_throw_not_promise => {
        r#"async function main(){const o=[];for(let i=0;i<2;i++){try{await (async()=>{throw i;})();}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["0,1"]
    };

    async_await_promise_resolve_after_catch_setup => {
        r#"async function main(){let ok=false;try{await Promise.reject("no");}catch{ok=true;}console.log(await Promise.resolve(ok));}main();"#,
        ["true"]
    };

    async_try_catch_typeerror_filter => {
        r#"async function main(){try{await Promise.reject(new TypeError("t"));}catch(e){console.log(e instanceof TypeError?"yes":"no");}}main();"#,
        ["yes"]
    };

    async_try_catch_reraise_unknown => {
        r#"async function main(){try{try{await Promise.reject(new TypeError("t"));}catch(e){if(e instanceof RangeError)throw e;throw "filtered";}}catch(e){console.log(e);}}main();"#,
        ["filtered"]
    };

    async_for_await_custom_async_iterator_throw => {
        r#"async function main(){const o=[];const it={async next(){if(o.length)return{done:true};o.push("n");if(o.length===2)throw "n2";return{value:o.length,done:false};},[Symbol.asyncIterator](){return this;}};const r=[];try{for await(const v of it)r.push(v);}catch(e){r.push("e");}console.log(r.join(","));}main();"#,
        ["1,e"]
    };

    async_loop_accumulator_survives_catch => {
        r#"async function main(){let sum=0;for(let i=1;i<=3;i++){try{sum+=await Promise.resolve(i);}catch{sum=-1;}if(i===2)try{await Promise.reject(0);}catch{}}console.log(sum);}main();"#,
        ["6"]
    };

    async_await_in_if_else_branches => {
        r#"async function main(){let r="";if(await Promise.resolve(true)){try{await Promise.reject("if");}catch(e){r=e;}}else{r="else";}console.log(r);}main();"#,
        ["if"]
    };

    async_await_in_while_condition => {
        r#"async function main(){let n=0;const o=[];while(await Promise.resolve(n<2)){o.push(n);n++;}console.log(o.join(","));}main();"#,
        ["0,1"]
    };

    async_for_await_with_break => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield 2;yield 3;}for await(const v of g()){o.push(v);if(v===2)break;}console.log(o.join(","));}main();"#,
        ["1,2"]
    };

    async_for_await_with_continue => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield 2;yield 3;}for await(const v of g()){if(v===2)continue;o.push(v);}console.log(o.join(","));}main();"#,
        ["1,3"]
    };

    async_function_expression_await_catch => {
        r#"async function main(){const f=async function(){try{await Promise.reject("fe");}catch(e){return e;}};console.log(await f());}main();"#,
        ["fe"]
    };

    async_arrow_await_catch => {
        r#"async function main(){const f=async()=>{try{await Promise.reject("ar");}catch(e){return e;}};console.log(await f());}main();"#,
        ["ar"]
    };

    async_method_shorthand_await_error => {
        r#"async function main(){const o={async m(){try{await Promise.reject("ms");}catch(e){return e;}}};console.log(await o.m());}main();"#,
        ["ms"]
    };

    async_try_with_multiple_await_before_throw => {
        r#"async function main(){try{await Promise.resolve(1);await Promise.resolve(2);await Promise.reject("third");}catch(e){console.log(e);}}main();"#,
        ["third"]
    };

    async_catch_binds_await_result => {
        r#"async function main(){try{throw "s";}catch(e){console.log(await Promise.resolve("c:"+e));}}main();"#,
        ["c:s"]
    };

    async_for_loop_empty_body_catch => {
        r#"async function main(){let c=0;for(let i=0;i<3;i++){try{await Promise.reject(i);}catch{c++;}}console.log(c);}main();"#,
        ["3"]
    };

    async_await_reject_in_try_after_resolve => {
        r#"async function main(){try{const a=await Promise.resolve(1);const b=await Promise.reject("after:"+a);console.log(b);}catch(e){console.log(e);}}main();"#,
        ["after:1"]
    };

    async_parallel_map_allSettled_errors => {
        r#"async function main(){const r=await Promise.allSettled([1,2].map(async x=>{if(x===2)throw "m"+x;return x;}));console.log(r[1].status+":"+r[1].reason);}main();"#,
        ["rejected:m2"]
    };

    async_for_await_from_promise_all => {
        r#"async function main(){const o=[];for await(const v of Promise.all([Promise.resolve(1),Promise.resolve(2)]))o.push(...v);console.log(o.join(","));}main();"#,
        ["1,2"]
    };

    async_iife_with_parameter_await_error => {
        r#"async function main(){const r=await(async(x)=>{try{await Promise.reject(x);}catch(e){return e;}})("param");console.log(r);}main();"#,
        ["param"]
    };

    async_nested_for_await_inner_throw => {
        r#"async function main(){const o=[];async function* outer(){yield 1;yield inner();async function* inner(){throw "inner";}}try{for await(const v of outer())o.push(String(v));}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:inner"]
    };

    async_try_catch_finally_order_on_reject => {
        r#"async function main(){const o=[];try{await Promise.reject("r");}catch(e){o.push("c");}finally{o.push("f");}console.log(o.join(","));}main();"#,
        ["c,f"]
    };

    async_await_in_object_destructure => {
        r#"async function main(){try{const{err}=await Promise.resolve({err:Promise.reject("d")});await err;}catch(e){console.log(e);}}main();"#,
        ["d"]
    };

    async_for_of_entries_await_catch => {
        r#"async function main(){const o=[];for(const[k,v]of Object.entries({a:1})){try{await Promise.reject(k+v);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["a1"]
    };

    async_while_await_increment_in_catch => {
        r#"async function main(){let i=0,o=[];while(i<2){try{await Promise.reject(i);}catch(e){o.push(e);i++;}}console.log(o.join(","));}main();"#,
        ["0,1"]
    };

    async_do_while_false_body_still_runs_once => {
        r#"async function main(){let ran=false;do{try{await Promise.reject("once");}catch{ran=true;}}while(false);console.log(ran);}main();"#,
        ["true"]
    };

    async_for_await_sync_throw_in_generator => {
        r#"async function main(){const o=[];async function* g(){yield 1;(()=>{throw "sync";})();}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:sync"]
    };

    async_await_error_in_callback_to_promise => {
        r#"async function main(){try{await new Promise((_,r)=>{queueMicrotask(()=>r("cb"));});}catch(e){console.log(e);}}main();"#,
        ["cb"]
    };

    async_retry_loop_three_attempts => {
        r#"async function main(){let n=0,o=[];for(let i=0;i<3;i++){try{if(++n<3)await Promise.reject("try"+n);else o.push("ok");}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["try1,try2,ok"]
    };

    async_for_await_yield_reject_promise => {
        r#"async function main(){const o=[];async function* g(){yield Promise.reject("yr");}try{for await(const v of g())o.push(await v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["e:yr"]
    };

    async_catch_logs_and_continues_loop => {
        r#"async function main(){const o=[];for(let i=0;i<3;i++){try{await Promise.reject(i);}catch(e){o.push("e"+e);continue;}o.push("ok");}console.log(o.join(","));}main();"#,
        ["e0,e1,e2"]
    };

    async_await_throws_non_error_primitive => {
        r#"async function main(){try{await (async()=>{throw 42;})();}catch(e){console.log(e);}}main();"#,
        ["42"]
    };

    async_await_throws_undefined => {
        r#"async function main(){try{await (async()=>{throw undefined;})();}catch(e){console.log(e===undefined);}}main();"#,
        ["true"]
    };

    async_for_await_from_async_generator_delegate => {
        r#"async function main(){const o=[];async function* inner(){yield 1;throw "d";}async function* outer(){yield* inner();}try{for await(const v of outer())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:d"]
    };

    async_promise_all_with_async_mapper_errors => {
        r#"async function main(){try{await Promise.all([1,2,3].map(async x=>{if(x===2)throw "mx";return x;}));}catch(e){console.log(e);}}main();"#,
        ["mx"]
    };

    async_sequential_await_in_try_per_iteration => {
        r#"async function main(){const o=[];for(const x of["a","b"]){try{o.push(await Promise.resolve(x));}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["a,b"]
    };

    async_iife_return_after_caught_await => {
        r#"async function main(){const v=await(async()=>{try{await Promise.reject("x");return 1;}catch{return 2;}})();console.log(v);}main();"#,
        ["2"]
    };

    async_for_await_symbol_async_iterator => {
        r#"async function main(){const o=[];const it={[Symbol.asyncIterator]:async function*(){yield 5;throw "sym";}};try{for await(const v of it)o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["5,e:sym"]
    };

    async_nested_try_only_inner_catches => {
        r#"async function main(){try{try{await Promise.reject("in");}catch(e){console.log("inner:"+e);}}catch{console.log("outer");}}main();"#,
        ["inner:in"]
    };

    async_outer_catch_on_inner_reraise => {
        r#"async function main(){try{try{await Promise.reject("a");}catch(e){throw "b:"+e;}}catch(e){console.log(e);}}main();"#,
        ["b:a"]
    };

    async_for_loop_with_await_and_break_in_try => {
        r#"async function main(){const o=[];for(let i=0;i<5;i++){try{o.push(await Promise.resolve(i));if(i===2)break;}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["0,1,2"]
    };

    async_await_in_template_literal => {
        r#"async function main(){try{const msg=`err:${await Promise.reject("tpl")}`;console.log(msg);}catch(e){console.log(e);}}main();"#,
        ["tpl"]
    };

    async_for_await_reject_on_first_iteration => {
        r#"async function main(){const o=[];async function* g(){throw "first";yield 1;}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["e:first"]
    };

    async_parallel_await_both_resolve => {
        r#"async function main(){const[a,b]=await Promise.all([Promise.resolve(1),Promise.resolve(2)]);console.log(a+b);}main();"#,
        ["3"]
    };

    async_catch_with_await_promise_all_inside => {
        r#"async function main(){try{await Promise.reject("x");}catch{const r=await Promise.all([1,2]);console.log(r.join("+"));}}main();"#,
        ["1+2"]
    };

    async_for_await_over_async_iterable_class => {
        r#"async function main(){const o=[];class It{async*[Symbol.asyncIterator](){yield 1;throw "cls";}}try{for await(const v of new It())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:cls"]
    };

    async_while_await_false_condition_exits => {
        r#"async function main(){let i=0;while(await Promise.resolve(i<0)){i++;}console.log(i);}main();"#,
        ["0"]
    };

    async_for_await_with_await_inside_loop_body => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield 2;}for await(const v of g()){o.push(await Promise.resolve(v*10));}console.log(o.join(","));}main();"#,
        ["10,20"]
    };

    async_iife_async_generator_inside => {
        r#"async function main(){const r=await(async()=>{async function* g(){yield 7;}let s=0;for await(const v of g())s+=v;return s;})();console.log(r);}main();"#,
        ["7"]
    };

    async_try_catch_with_await_in_both => {
        r#"async function main(){try{await Promise.reject("a");}catch(e){console.log(await Promise.resolve("c:"+e));}}main();"#,
        ["c:a"]
    };

    async_for_loop_decrement_with_catch => {
        r#"async function main(){const o=[];for(let i=2;i>=0;i--){try{await Promise.reject(i);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["2,1,0"]
    };

    async_await_reject_after_successful_awaits_in_loop => {
        r#"async function main(){const o=[];for(let i=0;i<3;i++){try{o.push(await Promise.resolve(i));if(i===2)await Promise.reject("done");}catch(e){o.push("e:"+e);}}console.log(o.join(","));}main();"#,
        ["0,1,2,e:done"]
    };

    async_for_await_multiple_consecutive_throws_not_reached => {
        r#"async function main(){const o=[];async function* g(){throw "only";throw "second";}try{for await(const v of g())o.push(v);}catch(e){o.push(e);}console.log(o.join(","));}main();"#,
        ["only"]
    };

    async_catch_empty_block_continues => {
        r#"async function main(){try{await Promise.reject("x");}catch{}console.log("after");}main();"#,
        ["after"]
    };

    async_for_await_recovers_and_continues_after_catch => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield Promise.reject("mid").catch(()=>"fixed");yield 3;}for await(const v of g())o.push(await v);console.log(o.join(","));}main();"#,
        ["1,fixed,3"]
    };

    async_iife_named_function_expression => {
        r#"async function main(){const r=await(async function named(){try{await Promise.reject("n");}catch(e){return e;}})();console.log(r);}main();"#,
        ["n"]
    };

    async_for_of_with_async_callback_pattern => {
        r#"async function main(){const o=[];const items=[1,2];for(const x of items){const v=await(async()=>{try{return await Promise.resolve(x*2);}catch{return -1;}})();o.push(v);}console.log(o.join(","));}main();"#,
        ["2,4"]
    };

    async_await_in_computed_property => {
        r#"async function main(){try{const o={[await Promise.reject("key")]:1};console.log(o);}catch(e){console.log(e);}}main();"#,
        ["key"]
    };

    async_for_await_over_rejecting_promise_directly => {
        r#"async function main(){const o=[];try{for await(const v of Promise.reject("direct"))o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["e:direct"]
    };

    async_loop_with_throw_not_in_promise => {
        r#"async function main(){const o=[];for(let i=0;i<2;i++){try{if(i===1)throw "raw";await Promise.resolve(i);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["raw"]
    };

    async_nested_await_three_levels_catch => {
        r#"async function main(){try{await(async()=>await(async()=>await Promise.reject("3l"))())();}catch(e){console.log(e);}}main();"#,
        ["3l"]
    };

    async_for_await_with_try_finally_no_catch => {
        r#"async function main(){const o=[];async function* g(){yield 1;throw "t";}try{for await(const v of g())o.push(v);}finally{o.push("f");}console.log(o.join(","));}main();"#,
        ["1,f"]
    };

    async_parallel_catch_only_one_branch_fails => {
        r#"async function main(){const o=await Promise.all([Promise.resolve(1),(async()=>{try{await Promise.reject("f");return 0;}catch{return 9;}})()]);console.log(o.join(","));}main();"#,
        ["1,9"]
    };

    async_for_loop_with_await_promise_resolve_chain => {
        r#"async function main(){const o=[];for(let i=0;i<2;i++){o.push(await Promise.resolve(i).then(x=>x+1));}console.log(o.join(","));}main();"#,
        ["1,2"]
    };

    async_iife_throw_after_await_resolve => {
        r#"async function main(){try{await(async()=>{await Promise.resolve();throw "post";})();}catch(e){console.log(e);}}main();"#,
        ["post"]
    };

    async_for_await_async_generator_with_await_inside => {
        r#"async function main(){const o=[];async function* g(){yield await Promise.resolve(1);throw "g";}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:g"]
    };

    async_catch_rethrows_after_await => {
        r#"async function main(){try{try{await Promise.reject("a");}catch(e){await Promise.resolve();throw e;}}catch(e){console.log(e);}}main();"#,
        ["a"]
    };

    async_for_await_empty_after_throw_not_run => {
        r#"async function main(){const o=[];async function* g(){throw "e";yield 1;}try{for await(const v of g())o.push(v);}catch(e){o.push("c:"+e);}console.log(o.join(","));}main();"#,
        ["c:e"]
    };

    async_while_with_await_throw_in_body => {
        r#"async function main(){let i=0,o=[];while(i<2){try{await(async()=>{if(i===1)throw "w";})();o.push(i);}catch(e){o.push(e);}i++;}console.log(o.join(","));}main();"#,
        ["0,w"]
    };

    async_try_nested_catch_different_errors => {
        r#"async function main(){try{try{await Promise.reject(new TypeError("t"));}catch(e){if(e instanceof TypeError)throw new RangeError("r");}}catch(e){console.log(e.name);}}main();"#,
        ["RangeError"]
    };

    async_for_await_from_array_of_async_iterables => {
        r#"async function main(){const o=[];async function* a(){yield 1;}async function* b(){throw "b";}for(const gen of[a(),b()]){try{for await(const v of gen)o.push(v);}catch(e){o.push("e:"+e);}}console.log(o.join(","));}main();"#,
        ["1,e:b"]
    };

    async_iife_with_finally_on_await_reject => {
        r#"async function main(){const o=[];await(async()=>{try{await Promise.reject("r");}catch(e){o.push(e);}finally{o.push("f");}})();console.log(o.join(","));}main();"#,
        ["r,f"]
    };

    async_for_loop_skip_iteration_via_continue_in_try => {
        r#"async function main(){const o=[];for(let i=0;i<3;i++){try{if(i===1)await Promise.reject("skip");o.push(i);}catch{continue;} }console.log(o.join(","));}main();"#,
        ["0,2"]
    };

    async_await_in_array_literal_element => {
        r#"async function main(){try{const a=[1,await Promise.reject("arr"),3];console.log(a);}catch(e){console.log(e);}}main();"#,
        ["arr"]
    };

    async_for_await_rejection_after_multiple_yields => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield 2;yield 3;throw "end";}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,2,3,e:end"]
    };

    async_catch_binds_error_message => {
        r#"async function main(){try{await Promise.reject(new Error("msg"));}catch(e){console.log(e.message);}}main();"#,
        ["msg"]
    };

    async_parallel_await_settled_both_reject => {
        r#"async function main(){const r=await Promise.allSettled([Promise.reject("a"),Promise.reject("b")]);console.log(r.map(x=>x.status).join(","));}main();"#,
        ["rejected,rejected"]
    };

    async_for_await_with_labeled_continue => {
        r#"async function main(){const o=[];outer:async function* g(){yield 1;yield 2;}for await(const v of g()){if(v===1)continue outer;o.push(v);}console.log(o.join(","));}main();"#,
        ["2"]
    };

    async_iife_reject_in_return_statement => {
        r#"async function main(){try{await(async()=>Promise.reject("ret"))();}catch(e){console.log(e);}}main();"#,
        ["ret"]
    };

    async_for_loop_await_with_outer_catch => {
        r#"async function main(){const o=[];try{for(let i=0;i<2;i++){await Promise.reject(i);}}catch(e){o.push("outer:"+e);}console.log(o.join(","));}main();"#,
        ["outer:0"]
    };

    async_for_await_generator_await_reject_inside => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield await Promise.reject("inner");}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:inner"]
    };

    async_nested_loops_inner_catch => {
        r#"async function main(){const o=[];for(let i=0;i<2;i++){for(let j=0;j<2;j++){try{await Promise.reject(i+""+j);}catch(e){o.push(e);}}}console.log(o.join(","));}main();"#,
        ["00,01,10,11"]
    };

    async_await_error_in_object_method => {
        r#"async function main(){const o={async fail(){await Promise.reject("om");}};try{await o.fail();}catch(e){console.log(e);}}main();"#,
        ["om"]
    };

    async_for_await_with_manual_iterator_close => {
        r#"async function main(){const o=[];async function* g(){try{yield 1;throw "x";}finally{o.push("cl");}}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,cl,e:x"]
    };

    async_iife_with_multiple_awaits_before_catch => {
        r#"async function main(){const r=await(async()=>{await Promise.resolve(1);await Promise.resolve(2);try{await Promise.reject("3");}catch{return "ok";}})();console.log(r);}main();"#,
        ["ok"]
    };

    async_for_await_reject_from_returned_promise => {
        r#"async function main(){const o=[];async function* g(){yield Promise.resolve(1);yield Promise.reject("rp");}try{for await(const v of g())o.push(await v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:rp"]
    };

    async_try_await_in_finally_after_catch_throw => {
        r#"async function main(){try{try{await Promise.reject("a");}catch(e){throw e;}}finally{console.log(await Promise.resolve("f"));} }main();"#,
        ["f"]
    };

    async_for_loop_await_with_switch_break => {
        r#"async function main(){const o=[];for(let i=0;i<3;i++){switch(i){case 0:try{await Promise.reject("s");}catch(e){o.push(e);}break;case 1:o.push("m");break;}}console.log(o.join(","));}main();"#,
        ["s,m"]
    };

    async_catch_with_conditional_reraise => {
        r#"async function main(){try{await Promise.reject("code:404");}catch(e){if(String(e).includes("404"))console.log("handled");else throw e;}}main();"#,
        ["handled"]
    };

    async_for_await_over_values_not_promises => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield 2;}for await(const v of g()){try{if(v===2)throw "v2";o.push(v);}catch(e){o.push("e:"+e);}}console.log(o.join(","));}main();"#,
        ["1,e:v2"]
    };

    async_iife_catch_typeerror_only => {
        r#"async function main(){const r=await(async()=>{try{await Promise.reject(new TypeError("t"));}catch(e){return e instanceof TypeError?"typed":"other";}})();console.log(r);}main();"#,
        ["typed"]
    };

    async_parallel_await_first_resolves_second_rejects => {
        r#"async function main(){const o=[];try{const[r]=await Promise.all([Promise.resolve("ok"),Promise.reject("bad")]);o.push(r);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["e:bad"]
    };

    async_for_await_with_spread_values => {
        r#"async function main(){const o=[];async function* g(){yield...[1,2];throw "s";}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,2,e:s"]
    };

    async_while_catch_increments_until_success => {
        r#"async function main(){let n=0;while(n<3){try{if(n<2)await Promise.reject(n);else break;}catch{n++;}}console.log(n);}main();"#,
        ["2"]
    };

    async_await_in_logical_not => {
        r#"async function main(){const v=!(await Promise.resolve(false));console.log(v);}main();"#,
        ["true"]
    };

    async_for_await_on_async_function_return => {
        r#"async function main(){const o=[];async function* src(){yield 9;}async function wrap(){return src();}for await(const v of await wrap())o.push(v);console.log(o.join(","));}main();"#,
        ["9"]
    };

    async_iife_error_in_promise_constructor => {
        r#"async function main(){try{await(async()=>await new Promise((_,r)=>r("ctor")))())();}catch(e){console.log(e);}}main();"#,
        ["ctor"]
    };

    async_for_loop_with_await_race_errors => {
        r#"async function main(){const o=[];for(let i=0;i<2;i++){try{await Promise.race([Promise.reject("r"+i),Promise.resolve(i)]);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["r0,r1"]
    };

    async_catch_return_from_async_loop => {
        r#"async function main(){async function run(){for(let i=0;i<3;i++){try{await Promise.reject(i);}catch(e){if(e===1)return "stop";}}return "end";}console.log(await run());}main();"#,
        ["stop"]
    };

    async_for_await_with_null_yield_skipped => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield null;yield 2;}for await(const v of g())o.push(v===null?"n":v);console.log(o.join(","));}main();"#,
        ["1,n,2"]
    };

    async_nested_iife_catch_at_outer => {
        r#"async function main(){try{await(async()=>{await(async()=>{await Promise.reject("inner");})();})();}catch(e){console.log(e);}}main();"#,
        ["inner"]
    };

    async_for_await_throw_after_await_in_generator => {
        r#"async function main(){const o=[];async function* g(){const x=await Promise.resolve(1);yield x;throw "after:"+x;}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,e:after:1"]
    };

    async_try_with_await_and_sync_throw_mix => {
        r#"async function main(){try{await Promise.resolve();throw "sync";}catch(e){console.log(e);}}main();"#,
        ["sync"]
    };

    async_for_loop_await_promise_all_per_item => {
        r#"async function main(){const o=[];for(const x of[1,2]){try{const[v]=await Promise.all([Promise.resolve(x)]);o.push(v);}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["1,2"]
    };

    async_iife_with_await_in_conditional_return => {
        r#"async function main(){const r=await(async()=>{try{await Promise.reject("x");}catch(e){return await Promise.resolve("r:"+e);}})();console.log(r);}main();"#,
        ["r:x"]
    };

    async_for_await_reject_with_custom_error => {
        r#"async function main(){const o=[];class E extends Error{}async function* g(){yield 1;throw new E("custom");}try{for await(const v of g())o.push(v);}catch(e){o.push(e instanceof E?"e:custom":"other");}console.log(o.join(","));}main();"#,
        ["1,e:custom"]
    };

    async_while_await_with_throw_in_finally => {
        r#"async function main(){const o=[];try{while(await Promise.resolve(false)){}}catch{}finally{try{await Promise.reject("wf");}catch(e){o.push(e);}}console.log(o.join(","));}main();"#,
        ["wf"]
    };

    async_parallel_for_await_two_generators => {
        r#"async function main(){const o=[];async function* a(){yield "a";}async function* b(){throw "b";}const run=async(g,l)=>{try{for await(const v of g())o.push(l+v);}catch(e){o.push(l+"e:"+e);}};await Promise.all([run(a(),""),run(b(),"")]);console.log(o.sort().join(","));}main();"#,
        ["a,b e:b"]
    };

    async_catch_on_awaited_async_throw => {
        r#"async function main(){try{await(async function(){throw new SyntaxError("s");})();}catch(e){console.log(e.name);}}main();"#,
        ["SyntaxError"]
    };

    async_for_await_over_single_rejecting_yield => {
        r#"async function main(){const o=[];async function* g(){yield Promise.reject("only");}try{for await(const v of g())o.push(await v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["e:only"]
    };

    async_iife_nested_try_finally_on_error => {
        r#"async function main(){const o=[];await(async()=>{try{try{await Promise.reject("a");}catch(e){o.push(e);}finally{o.push("f");}})();console.log(o.join(","));}main();"#,
        ["a,f"]
    };

    async_for_loop_with_await_and_optional_catch_binding => {
        r#"async function main(){const o=[];for(let i=0;i<2;i++){try{await Promise.reject(i);}catch{o.push("c"+i);}}console.log(o.join(","));}main();"#,
        ["c0,c1"]
    };

    async_for_await_with_break_in_catch => {
        r#"async function main(){const o=[];async function* g(){yield 1;yield 2;yield 3;}for await(const v of g()){try{if(v===2)throw "stop";o.push(v);}catch{break;}}console.log(o.join(","));}main();"#,
        ["1"]
    };

    async_await_reject_in_expression_statement => {
        r#"async function main(){try{await Promise.reject("stmt");}catch(e){console.log(e);}}main();"#,
        ["stmt"]
    };

    async_for_await_generator_finally_before_throw => {
        r#"async function main(){const o=[];async function* g(){try{yield 1;}finally{o.push("gf");}throw "after";}try{for await(const v of g())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();"#,
        ["1,gf,e:after"]
    };

}
