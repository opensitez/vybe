//! `ipairs` and `pairs` iteration traversal semantics (Lua 5.x §6.1)

lua_print! {
ipairs_stops_at_nil => {
    "local t = {10, 20, nil, 40}\nlocal n = 0\nfor _ in ipairs(t) do n = n + 1 end\nprint(n)\n",
    "2"
},
ipairs_empty => {
    "local n = 0\nfor _ in ipairs({}) do n = n + 1 end\nprint(n)\n",
    "0"
},
ipairs_one_based_indices => {
    "local first_i\nfor i, _ in ipairs({\"a\", \"b\"}) do\n  if not first_i then first_i = i end\nend\nprint(first_i)\n",
    "1"
},
ipairs_sum => {
    "local sum = 0\nfor _, v in ipairs({5, 10, 15}) do sum = sum + v end\nprint(sum)\n",
    "30"
},
ipairs_ignores_hash_keys => {
    "local t = {1, 2, x=99}\nlocal n = 0\nfor _ in ipairs(t) do n = n + 1 end\nprint(n)\n",
    "2"
},
ipairs_single => {
    "for i, v in ipairs({42}) do print(i .. \"=\" .. v) end\n",
    "1=42"
},
ipairs_sequence => {
    "local s = \"\"\nfor i, _ in ipairs({\"a\", \"b\", \"c\"}) do s = s .. i end\nprint(s)\n",
    "123"
},
pairs_visits_all_keys => {
    "local t = {1, 2, x=10}\nlocal n = 0\nfor _ in pairs(t) do n = n + 1 end\nprint(n)\n",
    "3"
},
pairs_empty => {
    "local n = 0\nfor _ in pairs({}) do n = n + 1 end\nprint(n)\n",
    "0"
},
pairs_string_keys => {
    "local t = {a=1, b=2, c=3}\nlocal n = 0\nfor _ in pairs(t) do n = n + 1 end\nprint(n)\n",
    "3"
},
ipairs_iterator_tuple => {
    "local it, s, i = ipairs({10, 20, 30})\nlocal _, v = it(s, i)\nprint(v)\n",
    "10"
} }
