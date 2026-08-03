//! Nested xpcall error recovery and error handler interactions (Lua 5.x §6.1)

lua_print! {
    nested_xpcall_inner_handles => {
        "local inner_run = false\nlocal outer_run = false\nlocal function inner_handler(e) inner_run = true; return \"inner:\"..e end\nlocal function outer_handler(e) outer_run = true; return \"outer:\"..e end\nlocal ok, val = xpcall(function()\n  return xpcall(function() error(\"fail\", 0) end, inner_handler)\nend, outer_handler)\nprint(ok, val, inner_run, outer_run)\n",
        "true false true false"
    },
    xpcall_handler_fails_nested => {
        "local function bad_handler(e) error(\"handler_broke\", 0) end\nlocal function outer_handler(e) return \"outer:\"..e end\nlocal ok, val = xpcall(function()\n  return xpcall(function() error(\"original\", 0) end, bad_handler)\nend, outer_handler)\nprint(ok, val)\n",
        "true false"
    },
    xpcall_returns_multi => {
        "local ok, a, b = xpcall(function() return \"a\", \"b\" end, function(e) return e end)\nprint(ok, a, b)\n",
        "true a b"
    },
    xpcall_handler_non_string => {
        "local ok, err = xpcall(function() error(\"boom\") end, function() return {code=500} end)\nprint(ok, type(err), err.code)\n",
        "false table 500"
    } }
