lua_print! {
    test_gc_count => { "local c1 = collectgarbage('count'); collectgarbage('collect'); local c2 = collectgarbage('count'); print(type(c1) == 'number' and type(c2) == 'number')", "true" },
    test_gc_collect => { "local ok = pcall(function() collectgarbage('collect') end); print(tostring(ok))", "true" },
    test_gc_stop_restart => { "collectgarbage('stop'); collectgarbage('restart'); print('ok')", "ok" },
    test_gc_step => { "local b = collectgarbage('step'); print(type(b) == 'boolean')", "true" },
    test_gc_setpause => { "local p = collectgarbage('setpause', 200); print(type(p) == 'number')", "true" },
    test_gc_setstepmul => { "local m = collectgarbage('setstepmul', 200); print(type(m) == 'number')", "true" },
    test_gc_isrunning => { "local r = collectgarbage('isrunning'); print(type(r) == 'boolean')", "true" },
    test_gc_invalid_opt => { "local ok = pcall(function() collectgarbage('invalid') end); print(tostring(ok))", "false" }
}
