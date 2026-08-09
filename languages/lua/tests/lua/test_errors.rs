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
    "true,fail"
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
xpcall_passes_extra_args_to_target => {
    "local function f(x, y) if x == y then error(\"equal\") end end\nlocal ok, msg = xpcall(f, function(e) return \"handled:\"..e end, 5, 5)\nprint(msg)\n",
    "handled:equal"
},
xpcall_handler_errors_themselves => {
    "local ok, msg = xpcall(function() error(\"first\") end, function(e) error(\"second\") end)\nprint(tostring(ok) .. \",\" .. tostring(msg:match(\"second\") ~= nil))\n",
    "false,false"
},
assert_returns_all_passed_arguments_on_success => {
    "local a, b, c = assert(10, \"err\", 30)\nprint(a .. \",\" .. tostring(b) .. \",\" .. tostring(c))\n",
    "10,err,30"
},
error_with_table_object => {
    "local err_obj = {code = 404}\nlocal ok, caught = pcall(function() error(err_obj) end)\nprint(ok == false and caught == err_obj and caught.code == 404)\n",
    "true"
},
xpcall_non_function_handler_raises_error => {
    "local ok, err = pcall(function() xpcall(function() error(\"fail\") end, \"not_a_function\") end)\nprint(ok)\n",
    "false"
},
pcall_nil_function_raises_error => {
    "local ok, err = pcall(pcall, nil)\nprint(ok)\n",
    "false"
},
xpcall_yield_inside_function => {
    "local co = coroutine.create(function()\n  local ok, res = xpcall(function() return coroutine.yield(\"yielding\") end, function(e) return e end)\n  return ok\nend)\nlocal _, val = coroutine.resume(co)\nlocal _, val2 = coroutine.resume(co)\nprint(val .. \" \" .. tostring(val2))\n",
    "yielding true"
},
error_with_level_two_removes_current_frame => {
    "local function f() error(\"my_err\", 2) end\nlocal ok, msg = pcall(f)\nprint(type(msg) == \"string\")\n",
    "true"
} }
