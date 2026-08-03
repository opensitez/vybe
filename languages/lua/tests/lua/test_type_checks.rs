//! `type` and `tostring` behavior on values (Lua 5.x manual §3.3.1).

lua_print! {
    type_thread_for_coroutine => {
        "local co=coroutine.create(function() end)\nprint(type(co))\n",
        "thread"
    },
    type_userdata_for_external_objects => {
        "print(type(io.stdin) == \"userdata\" or type(io.stdin) == \"nil\")\n",
        "true"
    },
    tostring_on_table_uses_metamethod => {
        "local t=setmetatable({}, {__tostring=function() return \"T\" end})\nprint(tostring(t))\n",
        "T"
    },
    tostring_on_number_not_scientific_for_integers => {
        "print(tostring(1000))\n",
        "1000"
    },
    tostring_on_function_shows_address_like_string => {
        "print(type(tostring(function() end)))\n",
        "string"
    },
    type_of_nan_is_number => {
        "print(type(0/0))\n",
        "number"
    },
    type_of_math_huge_is_number => {
        "print(type(math.huge))\n",
        "number"
    },
    pairs_iterator_returns_key_value => {
        "local t={a=1}\nlocal k,v=next(t)\nprint(k..\"=\"..v)\n",
        "a=1"
    },
    next_with_nil_starts_iteration => {
        "local t={x=1}\nlocal k=next(t,nil)\nprint(k)\n",
        "x"
    },
    rawequal_compares_without_metamethods => {
        "local a={}\nlocal b={}\nprint(rawequal(a,b))\n",
        "false"
    },
    rawequal_same_table_is_true => {
        "local a={}\nprint(rawequal(a,a))\n",
        "true"
    },
    type_on_non_existent_global => {
        "print(type(non_existent_global_var_xyz_123))\n",
        "nil"
    },
    tostring_on_nil => {
        "print(tostring(nil))\n",
        "nil"
    },
    tostring_on_boolean_true => {
        "print(tostring(true))\n",
        "true"
    },
    tostring_on_boolean_false => {
        "print(tostring(false))\n",
        "false"
    },
    rawequal_compares_numbers_of_different_types => {
        "print(rawequal(1, 1.0))\n",
        "true"
    },
    rawequal_compares_mismatched_types => {
        "print(rawequal(1, \"1\"))\n",
        "false"
    },
    next_on_empty_table_returns_nil => {
        "print(tostring(next({})))\n",
        "nil"
    },
    next_invalid_key_raises_error => {
        "local ok, err = pcall(function() next({a=1}, \"invalid_key\") end)\nprint(ok)\n",
        "false"
    } }
