lua_print! {
    test_sort_custom_basic => { "local t={3,1,2}; table.sort(t, function(a,b) return a>b end); print(t[1]..t[2]..t[3])", "321" },
    test_sort_custom_stable => { "local t={{a=1,b=2},{a=1,b=1}}; table.sort(t, function(x,y) return x.a<y.a end); print(t[1].b..t[2].b)", "21" },
    test_sort_custom_string => { "local t={'c','a','b'}; table.sort(t, function(a,b) return a<b end); print(t[1]..t[2]..t[3])", "abc" },
    test_sort_invalid_comp => { "local t={1,2}; local ok = pcall(function() table.sort(t, 'not a function') end); print(tostring(ok))", "false" },
    test_sort_comp_error => { "local t={1,2}; local ok = pcall(function() table.sort(t, function() error('boom') end) end); print(tostring(ok))", "false" },
    test_sort_comp_invalid_result => { "local t={1,2}; table.sort(t, function() return nil end); print(t[1]..t[2])", "12" }
}
