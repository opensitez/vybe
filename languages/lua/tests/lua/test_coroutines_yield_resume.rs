lua_print! {
    test_yield_basic => { "local co = coroutine.create(function() coroutine.yield(1) return 2 end); local ok, v1 = coroutine.resume(co); local ok2, v2 = coroutine.resume(co); print(v1..' '..v2)", "1 2" },
    test_yield_receives_resume_args => { "local co = coroutine.create(function() local x = coroutine.yield(); return x end); coroutine.resume(co); local ok, res = coroutine.resume(co, 42); print(res)", "42" },
    test_yield_multiple_values => { "local co = coroutine.create(function() coroutine.yield(1,2) end); local ok, a, b = coroutine.resume(co); print(a..' '..b)", "1 2" },
    test_yield_resume_multiple_values => { "local co = coroutine.create(function() local a, b = coroutine.yield(); return a+b end); coroutine.resume(co); local ok, res = coroutine.resume(co, 10, 20); print(res)", "30" },
    test_yield_inside_function => { "local function f() coroutine.yield(99) end; local co = coroutine.create(function() f() return 100 end); local ok, v1 = coroutine.resume(co); local ok2, v2 = coroutine.resume(co); print(v1..' '..v2)", "99 100" }
}
