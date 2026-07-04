lua_print! {
    test_concat_table_table => { "local mt={__concat=function(a,b) return a.v..b.v end}; local t1=setmetatable({v='a'}, mt); local t2=setmetatable({v='b'}, mt); print(t1..t2)", "ab" },
    test_concat_table_string => { "local mt={__concat=function(a,b) return a.v..b end}; local t1=setmetatable({v='a'}, mt); print(t1..'b')", "ab" },
    test_concat_string_table => { "local mt={__concat=function(a,b) return a..b.v end}; local t2=setmetatable({v='b'}, mt); print('a'..t2)", "ab" },
    test_concat_table_number => { "local mt={__concat=function(a,b) return a.v..tostring(b) end}; local t1=setmetatable({v='a'}, mt); print(t1..42)", "a42" },
    test_concat_chain => { "local mt={__concat=function(a,b) local av = type(a)=='table' and a.v or a; local bv = type(b)=='table' and b.v or b; return av..bv end}; local t1=setmetatable({v='1'}, mt); local t2=setmetatable({v='2'}, mt); print(t1..t2..t1)", "121" },
    test_concat_fallback_error => { "local t1={}; local t2={}; local ok = pcall(function() return t1..t2 end); print(ok)", "false" }
}
