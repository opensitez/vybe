lua_print! {
    test_weak_keys_basic => { "local t=setmetatable({}, {__mode='k'}); local k={}; t[k]=1; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_keys_strong_values => { "local t=setmetatable({}, {__mode='k'}); local k={}; local v={}; t[k]=v; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_keys_string_keys_are_strong => { "local t=setmetatable({}, {__mode='k'}); local k='str'; t[k]=1; k=nil; collectgarbage(); print(t.str)", "1" },
    test_weak_keys_number_keys_are_strong => { "local t=setmetatable({}, {__mode='k'}); local k=42; t[k]=1; k=nil; collectgarbage(); print(t[42])", "1" },
    test_weak_keys_boolean_keys_are_strong => { "local t=setmetatable({}, {__mode='k'}); local k=true; t[k]=1; k=nil; collectgarbage(); print(t[true])", "1" },
    test_weak_keys_function_keys => { "local t=setmetatable({}, {__mode='k'}); local k=function() end; t[k]=1; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_keys_thread_keys => { "local t=setmetatable({}, {__mode='k'}); local k=coroutine.create(function() end); t[k]=1; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_keys_re_add => { "local t=setmetatable({}, {__mode='k'}); local k={}; t[k]=1; k=nil; collectgarbage(); local k2={}; t[k2]=2; local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "1" },
    test_weak_keys_mixed_strong_weak => { "local t=setmetatable({}, {__mode='k'}); local k1={}; local k2={}; t[k1]=1; t[k2]=2; k1=nil; collectgarbage(); print((t[k1] or 'nil')..' '..t[k2])", "nil 2" },
    test_weak_keys_ephemeron_behavior => { "local t=setmetatable({}, {__mode='k'}); local k={}; t[k]=k; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" }
}
