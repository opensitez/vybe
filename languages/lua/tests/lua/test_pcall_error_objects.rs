//! `pcall` error recovery — multiple return values, object errors, nested pcall (Lua 5.x §2.4, §6.1)

lua_print! {
    pcall_multi_return_success => {
        "local ok, a, b = pcall(function() return 1, 2 end)\nprint(ok, a, b)\n",
        "true\t1\t2"
    },
    pcall_error_msg => {
        "local ok, msg = pcall(function() error(\"oops\") end)\nprint(ok, type(msg))\n",
        "false\tstring"
    },
    pcall_error_table_object => {
        "local ok, err = pcall(function() error({code=42}) end)\nprint(ok, err.code)\n",
        "false\t42"
    },
    pcall_error_int_object => {
        "local ok, err = pcall(function() error(99) end)\nprint(ok, err)\n",
        "false\t99"
    },
    nested_pcall_inner => {
        "local outer_ok = pcall(function()\n  local inner_ok = pcall(function() error(\"inner\") end)\n  print(inner_ok)\nend)\nprint(outer_ok)\n",
        "false\ntrue"
    },
    pcall_pass_args => {
        "local ok, v = pcall(function(x) return x * 2 end, 21)\nprint(ok, v)\n",
        "true\t42"
    },
    pcall_runtime_error_catch => {
        "local ok, _ = pcall(function() return ({}):missing() end)\nprint(ok)\n",
        "false"
    },
    pcall_level_one_loc => {
        "local ok, msg = pcall(function() error(\"fail\", 1) end)\nprint(ok, type(msg))\n",
        "false\tstring"
    },
    pcall_level_zero_raw => {
        "local ok, msg = pcall(function() error(\"raw\", 0) end)\nprint(ok, msg)\n",
        "false\traw"
    },
    pcall_normal_execution => {
        "local ok = pcall(function()\n  local x = 1 + 1\n  _ = x\nend)\nprint(ok)\n",
        "true"
    },
    pcall_table_fields => {
        "local ok, e = pcall(function() error({a=1, b=2}) end)\nprint(e.a + e.b)\n",
        "3"
    },
}
