//! Table keys traversal behavior during insertions and deletions (Lua 5.x §6.1, §6.6)

lua_print! {
    iter_empty_table => {
        "local count = 0\nfor k, v in pairs({}) do count = count + 1 end\nprint(count)\n",
        "0"
    },
    iter_nil_start_next => {
        "local t = {a=1}\nlocal k, v = next(t, nil)\nprint(k, v)\n",
        "a\t1"
    },
    iter_delete_current => {
        "local t = {a=1, b=2, c=3}\nlocal keys = {}\nfor k, v in pairs(t) do\n  keys[#keys+1] = k\n  t[k] = nil\nend\nprint(#keys)\n",
        "3"
    },
    iter_ipairs_stops_nil => {
        "local t = {10, 20, nil, 40}\nlocal values = {}\nfor i, v in ipairs(t) do values[#values+1] = v end\nprint(table.concat(values, \",\"))\n",
        "10,20"
    },
    iter_pairs_all_hash => {
        "local t = {x=1, y=2}\nlocal count = 0\nfor k, v in pairs(t) do count = count + 1 end\nprint(count)\n",
        "2"
    },
}
