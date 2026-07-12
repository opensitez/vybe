//! `xpcall` message handler — error intercepting and modification (Lua 5.x §6.1)

lua_print! {
    xpcall_handler_receives_err => {
        "local function handler(err) return \"caught:\" .. tostring(err) end\nlocal ok, msg = xpcall(function() error(\"boom\") end, handler)\nprint(ok, msg)\n",
        "false\tcaught:input:1: boom"
    },
    xpcall_handler_table => {
        "local function handler(err) return {msg=err} end\nlocal ok, r = xpcall(function() error(\"fail\", 0) end, handler)\nprint(ok, r.msg)\n",
        "false\tfail"
    },
    xpcall_success_val => {
        "local ok, v = xpcall(function() return 42 end, function(e) return e end)\nprint(ok, v)\n",
        "true\t42"
    },
    xpcall_raw_err => {
        "local handler_got = nil\nxpcall(function() error(\"raw\", 0) end, function(e) handler_got = e end)\nprint(handler_got)\n",
        "raw"
    },
    xpcall_forward_args => {
        "local ok, v = xpcall(function(a, b) return a + b end, function(e) return e end, 10, 32)\nprint(ok, v)\n",
        "true\t42"
    },
    xpcall_table_error => {
        "local function handler(e) return e.code end\nlocal ok, v = xpcall(function() error({code=7}) end, handler)\nprint(ok, v)\n",
        "false\t7"
    },
    xpcall_nested_capture => {
        "local inner_ok\nlocal outer_ok = xpcall(function()\n  inner_ok = xpcall(function() error(\"in\") end, function() return \"h\" end)\nend, function(e) return e end)\nprint(inner_ok, outer_ok)\n",
        "false\ttrue"
    },
    xpcall_no_error_no_handler => {
        "local called = false\nlocal ok = xpcall(function() return 1 end, function() called = true end)\nprint(ok, called)\n",
        "true\tfalse"
    },
}
