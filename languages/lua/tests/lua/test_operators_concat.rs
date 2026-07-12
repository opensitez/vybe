lua_print! {
    test_concat_string_string => { "print('abc' .. 'def')", "abcdef" },
    test_concat_string_number => { "print('abc' .. 10)", "abc10" },
    test_concat_number_string => { "print(10 .. 'abc')", "10abc" },
    test_concat_number_number => { "print(10 .. 20)", "1020" },
    test_concat_multiple => { "print('a' .. 'b' .. 'c')", "abc" },
    test_concat_right_associative => { "local t={}; setmetatable(t, {__concat=function(a,b) return tostring(a)..tostring(b) end}); print(type(t .. 'a' .. 'b'))", "string" },
    test_concat_right_assoc_eval => { "local c=0; local function f(n) c=c+1 return n end; local _ = f(1) .. f(2) .. f(3); print(c)", "3" },
    test_concat_invalid_boolean => { "local ok = pcall(function() return 'a' .. true end); print(tostring(ok))", "false" },
    test_concat_invalid_table => { "local ok = pcall(function() return 'a' .. {} end); print(tostring(ok))", "false" },
    test_concat_invalid_nil => { "local ok = pcall(function() return 'a' .. nil end); print(tostring(ok))", "false" }
}
