lua_print! {
    test_len_table => { "local mt={__len=function(t) return 42 end}; local t=setmetatable({1,2,3}, mt); print(#t)", "42" },
    test_len_string_metamethod => { "debug.setmetatable('', {__len=function(s) return 99 end}); print(#'abc')", "99" },
    test_len_fallback => { "local t={1,2,3}; print(#t)", "3" },
    test_len_error_no_metamethod_on_userdata => { "local t=io.stdin; local ok = pcall(function() return #t end); print(ok)", "false" }
}
