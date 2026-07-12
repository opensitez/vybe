lua_print! {
    test_coroutine_create => { "local co = coroutine.create(function() return 1 end); print(type(co))", "thread" },
    test_coroutine_status_suspended => { "local co = coroutine.create(function() end); print(coroutine.status(co))", "suspended" },
    test_coroutine_status_running => { "local st; local co = coroutine.create(function() st=coroutine.status(coroutine.running()) end); coroutine.resume(co); print(st)", "running" },
    test_coroutine_status_dead => { "local co = coroutine.create(function() end); coroutine.resume(co); print(coroutine.status(co))", "dead" },
    test_coroutine_status_normal => { "local st; local co1 = coroutine.create(function() st=coroutine.status(coroutine.running()) end); local co2 = coroutine.create(function() coroutine.resume(co1) end); coroutine.resume(co2); print(st)", "running" },
    test_coroutine_resume_basic => { "local co = coroutine.create(function() return 42 end); local ok, res = coroutine.resume(co); print(tostring(ok)..' '..res)", "true 42" },
    test_coroutine_resume_args => { "local co = coroutine.create(function(a,b) return a+b end); local ok, res = coroutine.resume(co, 10, 20); print(tostring(ok)..' '..res)", "true 30" }
}
