lua_print! {
    test_block_basic => { "local a=1; do local a=2 end; print(a)", "1" },
    test_block_multiple_shadowing => { "local a=1; do local a=2; do local a=3 end; print(a) end; print(a)", "2\\n1" },
    test_block_function_scope => { "local f; do local a=42; f = function() return a end end; print(f())", "42" },
    test_block_empty => { "do end; print('ok')", "ok" },
    test_block_return => { "local function f() do return 1 end; return 2 end; print(f())", "1" },
    test_block_break => { "local a=0; while true do do break end; a=1 end; print(a)", "0" },
    test_block_local_function => { "local f1; do local function f2() return 42 end; f1 = f2 end; print(f1())", "42" },
    test_block_local_function_shadowing => { "local function f() return 1 end; do local function f() return 2 end; print(f()) end; print(f())", "2\\n1" }
}
