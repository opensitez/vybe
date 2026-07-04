lua_print! {
    test_coroutine_isyieldable_main => { "print(tostring(coroutine.isyieldable()))", "false" },
    test_coroutine_isyieldable_coro => { "local co = coroutine.create(function() print(tostring(coroutine.isyieldable())) end); coroutine.resume(co)", "true" },
    test_coroutine_running_main => { "local co, is_main = coroutine.running(); print(type(co)..' '..tostring(is_main))", "thread true" },
    test_coroutine_running_coro => { "local co = coroutine.create(function() local c, m = coroutine.running(); print(type(c)..' '..tostring(m)) end); coroutine.resume(co)", "thread false" },
    test_coroutine_yield_no_args => { "local co = coroutine.create(function() coroutine.yield() return 1 end); coroutine.resume(co); local ok, r = coroutine.resume(co); print(r)", "1" },
    test_coroutine_error_in_yield => { "local co = coroutine.create(function() pcall(function() coroutine.yield() end) return 1 end); coroutine.resume(co); local ok, r = coroutine.resume(co); print(r)", "1" }
}
