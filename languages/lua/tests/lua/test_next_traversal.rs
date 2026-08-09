//! `next` function — raw table traversal (Lua 5.x §6.1)

lua_print! {
next_empty_nil => {
    "print(tostring(next({})))\n",
    "nil"
},
next_key_val_types => {
    "local t = {x=1}\nlocal k, v = next(t)\nprint(type(k) .. \":\" .. type(v))\n",
    "string:number"
},
next_nil_start => {
    "local t = {a=10}\nlocal k, v = next(t, nil)\nprint(k .. \"=\" .. v)\n",
    "a=10"
},
next_keys_count => {
    "local t = {a=1, b=2, c=3}\nlocal n = 0\nlocal k = nil\nrepeat\n  k = next(t, k)\n  if k then n = n + 1 end\nuntil not k\nprint(n)\n",
    "3"
},
next_past_last_nil => {
    "local t = {x=1}\nlocal k = next(t, nil)\nprint(tostring(next(t, k)))\n",
    "nil"
},
next_numeric_key => {
    "local t = {10, 20, 30}\nlocal k, v = next(t, nil)\nprint(type(k))\n",
    "number"
},
next_traversal_loop => {
    "local t = {p=1, q=2}\nlocal seen = 0\nlocal k = nil\nrepeat\n  k, _ = next(t, k)\n  if k then seen = seen + 1 end\nuntil not k\nprint(seen)\n",
    "2"
},
next_nonempty_check => {
    "local function nonempty(t) return next(t) ~= nil end\nprint(nonempty({1}))\n",
    "true"
},
next_empty_check => {
    "local function nonempty(t) return next(t) ~= nil end\nprint(nonempty({}))\n",
    "false"
} }
