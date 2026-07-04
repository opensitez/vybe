lua_print! {
    test_call_table => { "local t=setmetatable({}, {__call=function(tbl, a, b) return a+b end}); print(t(10, 20))", "30" },
    test_call_string => { "debug.setmetatable('', {__call=function(s, a) return s..a end}); print(('hello ')('world'))", "hello world" },
    test_call_passes_self => { "local target; local t=setmetatable({}, {__call=function(tbl) target=tbl end}); t(); print(tostring(t==target))", "true" },
    test_call_chain => { "local t1=setmetatable({}, {__call=function(tbl, a) return a*2 end}); local t2=setmetatable({}, {__call=t1}); print(t2(5))", "10" },
    test_call_error_no_metamethod => { "local t={}; local ok = pcall(function() t() end); print(ok)", "false" }
}
