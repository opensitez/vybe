lua_print! {
    test_debug_getinfo_function => { "local info = debug.getinfo(print); print(type(info))", "table" },
    test_debug_getinfo_level => { "local info = debug.getinfo(1); print(type(info))", "table" },
    test_debug_getinfo_invalid_level => { "local info = debug.getinfo(100); print(tostring(info))", "nil" },
    test_debug_getinfo_fields => { "local info = debug.getinfo(1, 'nlS'); print(type(info.name)..' '..type(info.currentline)..' '..type(info.short_src))", "nil number string" },
    test_debug_traceback_basic => { "local tb = debug.traceback(); print(type(tb))", "string" },
    test_debug_traceback_message => { "local tb = debug.traceback('my error'); print(tostring(string.find(tb, 'my error') ~= nil))", "true" },
    test_debug_traceback_level => { "local tb = debug.traceback('msg', 2); print(type(tb))", "string" }
}
