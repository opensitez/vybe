lua_print! {
    test_wrap_basic => { "local f = coroutine.wrap(function() return 42 end); print(f())", "42" },
    test_wrap_yield => { "local f = coroutine.wrap(function() coroutine.yield(1) return 2 end); print(f()..' '..f())", "1 2" },
    test_wrap_args => { "local f = coroutine.wrap(function(a,b) return a+b end); print(f(10, 20))", "30" },
    test_wrap_yield_args => { "local f = coroutine.wrap(function() local x = coroutine.yield(); return x end); f(); print(f(99))", "99" },
    test_wrap_error_throws => { "local f = coroutine.wrap(function() error('boom') end); local ok, err = pcall(f); print(tostring(ok))", "false" }
}
