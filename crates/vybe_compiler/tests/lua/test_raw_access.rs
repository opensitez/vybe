//! `rawget` / `rawset` / `rawequal` / `rawlen` — bypass metamethods (Lua 5.x §6.1)

lua_print! {
    rawget_reads_without_index => {
        "local t = setmetatable({}, {__index = function() return 99 end})\nrawset(t, \"x\", 42)\nprint(rawget(t, \"x\"))\n",
        "42"
    },
    rawget_missing_nil => {
        "local t = setmetatable({}, {__index = function() return 99 end})\nprint(tostring(rawget(t, \"missing\")))\n",
        "nil"
    },
    rawset_writes_without_newindex => {
        "local store = {}\nlocal t = setmetatable({}, {__newindex = function(_, k, v) store[k] = v end})\nrawset(t, \"k\", 7)\nprint(rawget(t, \"k\"))\n",
        "7"
    },
    rawset_no_newindex_trigger => {
        "local called = false\nlocal t = setmetatable({}, {__newindex = function() called = true end})\nrawset(t, \"x\", 1)\nprint(called)\n",
        "false"
    },
    rawequal_identity => {
        "local t = {}\nprint(rawequal(t, t))\n",
        "true"
    },
    rawequal_distinct => {
        "print(rawequal({}, {}))\n",
        "false"
    },
    rawequal_ignores_eq => {
        "local mt = {__eq = function() return true end}\nlocal a = setmetatable({}, mt)\nlocal b = setmetatable({}, mt)\nprint(rawequal(a, b))\n",
        "false"
    },
    rawequal_str_literal => {
        "print(rawequal(\"hello\", \"hello\"))\n",
        "true"
    },
    rawequal_mismatched_types => {
        "print(rawequal(1, \"1\"))\n",
        "false"
    },
    rawlen_array => {
        "local t = {10, 20, 30}\nprint(rawlen(t))\n",
        "3"
    },
    rawlen_ignores_len => {
        "local t = setmetatable({1, 2, 3}, {__len = function() return 999 end})\nprint(rawlen(t))\n",
        "3"
    },
    rawlen_str => {
        "print(rawlen(\"hello\"))\n",
        "5"
    },
    rawget_fn => {
        "local t = {}\nrawset(t, \"fn\", function() return 5 end)\nprint(rawget(t, \"fn\")())\n",
        "5"
    },
}
