//! Protected calls — `pcall`, `assert` (Lua 5.x manual §4.4).

lua_print! {
    pcall_returns_true_on_success => {
        "local ok = pcall(function() return 1 end)\nprint(ok)\n",
        "true"
    },
    pcall_returns_function_result_on_success => {
        "local _, v = pcall(function() return \"ok\" end)\nprint(v)\n",
        "ok"
    },
    pcall_catches_error_and_returns_false => {
        "local ok = pcall(function() error(\"boom\") end)\nprint(ok)\n",
        "false"
    },
    pcall_returns_error_object_as_second_value => {
        "local _, err = pcall(function() error(\"boom\") end)\nprint(err)\n",
        "boom"
    },
    assert_with_true_condition_returns_value => {
        "print(assert(true, \"fail\"))\n",
        "true"
    },
    assert_with_false_condition_raises => {
        "local ok = pcall(function() assert(false, \"bad\") end)\nprint(ok)\n",
        "false"
    },
    pcall_invokes_function_with_arguments => {
        "local _, v = pcall(function(x) return x * 2 end, 6)\nprint(v)\n",
        "12"
    },
    error_with_level_zero_still_abortable_by_pcall => {
        "local ok = pcall(function() error(\"stop\", 0) end)\nprint(ok)\n",
        "false"
    },
    xpcall_invokes_error_handler => {
        "local function handler(err) return \"handled:\" .. err end\nlocal ok, msg = xpcall(function() error(\"boom\") end, handler)\nprint(msg)\n",
        "handled:boom"
    },
    pcall_nested_error_bubbles_once => {
        "local ok = pcall(function()\n  pcall(function() error(\"inner\") end)\n  error(\"outer\")\nend)\nprint(ok)\n",
        "false"
    },
    error_function_stops_execution => {
        "local ok = pcall(function() error(\"stop\") end)\nprint(ok)\n",
        "false"
    },
    assert_false_triggers_error_message => {
        "local ok, msg = pcall(function() assert(false, \"bad\") end)\nprint(msg)\n",
        "bad"
    },
    pcall_on_builtin_tonumber_invalid => {
        "local ok, v = pcall(function() return tonumber(\"x\") end)\nprint(tostring(v))\n",
        "nil"
    },
    error_without_argument_defaults_message => {
        "local ok, msg = pcall(function() error() end)\nprint(type(msg))\n",
        "string"
    },
    assert_nil_condition_fails => {
        "local ok = pcall(function() assert(nil) end)\nprint(ok)\n",
        "false"
    },
    assert_number_zero_passes => {
        "print(assert(0))\n",
        "0"
    },
    pcall_catches_division_by_zero => {
        "local ok = pcall(function() return 1 / 0 end)\nprint(ok)\n",
        "true"
    },
    xpcall_returns_false_on_failure => {
        "local ok = xpcall(function() error(\"e\") end, function(e) return e end)\nprint(ok)\n",
        "false"
    },
    pcall_with_multiple_return_values => {
        "local ok, a, b = pcall(function() return 1, 2 end)\nprint(a + b)\n",
        "3"
    },
    error_second_argument_level_is_number => {
        "local ok = pcall(function() error(\"msg\", 1) end)\nprint(ok)\n",
        "false"
    },
}
