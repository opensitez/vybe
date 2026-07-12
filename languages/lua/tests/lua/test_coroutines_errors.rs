lua_print! {
    test_coroutine_error_propagation => { "local co = coroutine.create(function() error('boom') end); local ok, err = coroutine.resume(co); print(tostring(ok))", "false" },
    test_coroutine_error_dead => { "local co = coroutine.create(function() end); coroutine.resume(co); local ok, err = coroutine.resume(co); print(tostring(ok))", "false" },
    test_coroutine_resume_running => { "local co; co = coroutine.create(function() local ok, err = coroutine.resume(co); return tostring(ok) end); local ok, res = coroutine.resume(co); print(res)", "false" },
    test_coroutine_pcall_inside => { "local co = coroutine.create(function() local ok = pcall(function() error('x') end); return ok end); local ok, res = coroutine.resume(co); print(tostring(res))", "false" }
}
