lua_print! {
    test_metatable_protect => { "local t={}; setmetatable(t, {__metatable='protected'}); local ok = pcall(function() setmetatable(t, {}) end); print(tostring(ok))", "false" },
    test_metatable_protect_get => { "local t={}; setmetatable(t, {__metatable='protected'}); print(getmetatable(t))", "protected" },
    test_metatable_eq_same_mt => { "local mt={__eq=function(a,b) return a.x==b.x end}; local t1={x=1}; setmetatable(t1, mt); local t2={x=1}; setmetatable(t2, mt); print(tostring(t1==t2))", "true" },
    test_metatable_eq_diff_mt => { "local t1={x=1}; setmetatable(t1, {__eq=function(a,b) return a.x==b.x end}); local t2={x=1}; setmetatable(t2, {__eq=function(a,b) return a.x==b.x end}); print(tostring(t1==t2))", "false" },
    test_metatable_lt => { "local mt={__lt=function(a,b) return a.x<b.x end}; local t1={x=1}; setmetatable(t1, mt); local t2={x=2}; setmetatable(t2, mt); print(tostring(t1<t2))", "true" },
    test_metatable_le_fallback => { "local mt={__lt=function(a,b) return a.x<b.x end}; local t1={x=1}; setmetatable(t1, mt); local t2={x=2}; setmetatable(t2, mt); print(tostring(t1<=t2))", "true" }
}
