lua_print! {
    test_weak_values_basic => { "local t=setmetatable({}, {__mode='v'}); local v={}; t[1]=v; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_values_strong_keys => { "local t=setmetatable({}, {__mode='v'}); local k={}; local v={}; t[k]=v; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_values_string_values_are_strong => { "local t=setmetatable({}, {__mode='v'}); local v='str'; t[1]=v; v=nil; collectgarbage(); print(t[1])", "str" },
    test_weak_values_number_values_are_strong => { "local t=setmetatable({}, {__mode='v'}); local v=42; t[1]=v; v=nil; collectgarbage(); print(t[1])", "42" },
    test_weak_values_boolean_values_are_strong => { "local t=setmetatable({}, {__mode='v'}); local v=true; t[1]=v; v=nil; collectgarbage(); print(t[1])", "true" },
    test_weak_values_function_values => { "local t=setmetatable({}, {__mode='v'}); local v=function() end; t[1]=v; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_values_thread_values => { "local t=setmetatable({}, {__mode='v'}); local v=coroutine.create(function() end); t[1]=v; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_values_re_add => { "local t=setmetatable({}, {__mode='v'}); local v={}; t[1]=v; v=nil; collectgarbage(); local v2={}; t[2]=v2; local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "1" },
    test_weak_values_mixed_strong_weak => { "local t=setmetatable({}, {__mode='v'}); local v1={}; local v2={}; t[1]=v1; t[2]=v2; v1=nil; collectgarbage(); print((t[1] or 'nil')..' '..tostring(t[2]==v2))", "nil true" }
}
