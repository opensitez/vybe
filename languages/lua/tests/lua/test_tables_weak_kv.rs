lua_print! {
    test_weak_kv_key_collected => { "local t=setmetatable({}, {__mode='kv'}); local k={}; local v={}; t[k]=v; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_kv_value_collected => { "local t=setmetatable({}, {__mode='kv'}); local k={}; local v={}; t[k]=v; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_kv_both_collected => { "local t=setmetatable({}, {__mode='kv'}); local k={}; local v={}; t[k]=v; k=nil; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" },
    test_weak_kv_neither_collected => { "local t=setmetatable({}, {__mode='kv'}); local k={}; local v={}; t[k]=v; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "1" },
    test_weak_kv_ephemeron => { "local t=setmetatable({}, {__mode='kv'}); local k={}; local v={}; v.key=k; t[k]=v; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; print(cnt)", "0" }
}
