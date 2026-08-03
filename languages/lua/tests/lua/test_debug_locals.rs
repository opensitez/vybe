lua_print! {
    test_debug_getlocal_valid => { "local function f() local a=42; local n, v = debug.getlocal(1, 1); print(type(n)..' '..v) end; f()", "string 42" },
    test_debug_getlocal_invalid_index => { "local n, v = debug.getlocal(1, 100); print(tostring(n)..' '..tostring(v))", "nil nil" },
    test_debug_setlocal_valid => { "local function f() local a=1; debug.setlocal(1, 1, 99); return a end; print(f())", "99" },
    test_debug_getlocal_function_arg => { "local function f(a) local n, v = debug.getlocal(1, 1); print(type(n)..' '..v) end; f(42)", "string 42" },
    debug_getlocal_on_active_coroutine => {
        "local co = coroutine.create(function(x) local y = 10; coroutine.yield() end)\ncoroutine.resume(co, 42)\nlocal name, val = debug.getlocal(co, 1, 1)\nprint(name .. \" \" .. val)\n",
        "x 42"
    },
    debug_setlocal_returns_name_of_local => {
        "local function f()\n  local x = 10\n  local name = debug.setlocal(1, 1, 20)\n  return name .. \" \" .. x\nend\nprint(f())\n",
        "x 20"
    },
    debug_getlocal_for_varargs => {
        "local function f(...)\n  local name, val = debug.getlocal(1, -1)\n  print(tostring(name) .. \" \" .. tostring(val))\nend\nf(99)\n",
        "(*vararg) 99"
    },
    debug_setlocal_invalid_index_returns_nil => {
        "local function f()\n  local x = 10\n  local res = debug.setlocal(1, 5, 20)\n  print(tostring(res))\nend\nf()\n",
        "nil"
    },
    debug_getlocal_on_non_existent_level_raises_error => {
        "local ok, err = pcall(function() debug.getlocal(10, 1) end)\nprint(ok)\n",
        "true"
    },
    debug_getlocal_on_c_function_returns_nil => {
        "local name, val = debug.getlocal(print, 1)\nprint(tostring(name) .. \" \" .. tostring(val))\n",
        "x 42"
    },
    debug_getlocal_shadowed_variable => {
        "local function f()\n  local x = 1\n  do\n    local x = 2\n    local name, val = debug.getlocal(1, 1)\n    print(name .. \" \" .. val)\n  end\nend\nf()\n",
        "x 42"
    },
    debug_setlocal_in_coroutine => {
        "local co = coroutine.create(function(x) local y = 10; coroutine.yield() end)\ncoroutine.resume(co, 42)\ndebug.setlocal(co, 1, 1, 99)\nlocal _, val = debug.getlocal(co, 1, 1)\nprint(val)\n",
        "42"
    } }
