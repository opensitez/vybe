lua_print! {
    test_debug_gethook_empty => { "local h, m, c = debug.gethook(); print(tostring(h)..' '..tostring(m)..' '..tostring(c))", "nil  0" },
    test_debug_sethook_call => { "local c=0; debug.sethook(function() c=c+1 end, 'c'); local function f() end; f(); debug.sethook(); print(c)", "1" },
    test_debug_sethook_return => { "local c=0; debug.sethook(function() c=c+1 end, 'r'); local function f() end; f(); debug.sethook(); print(tostring(c > 0))", "true" },
    test_debug_sethook_line => { "local c=0; debug.sethook(function() c=c+1 end, 'l'); local x=1; local y=2; debug.sethook(); print(tostring(c > 0))", "true" },
    test_debug_sethook_count => { "local c=0; debug.sethook(function() c=c+1 end, '', 10); for i=1,20 do end; debug.sethook(); print(tostring(c > 0))", "true" }
}
