//! `table.unpack` with range arguments — selecting a sub-range (Lua 5.x §6.6)

lua_print! {
    unpack_full => {
        "local t = {10, 20, 30}\nprint(table.unpack(t))\n",
        "10\t20\t30"
    },
    unpack_start_index => {
        "local t = {10, 20, 30}\nprint(table.unpack(t, 2))\n",
        "20\t30"
    },
    unpack_start_and_end => {
        "local t = {10, 20, 30, 40}\nprint(table.unpack(t, 2, 3))\n",
        "20\t30"
    },
    unpack_single_element => {
        "local t = {5, 6, 7}\nprint(table.unpack(t, 2, 2))\n",
        "6"
    },
    unpack_to_args => {
        "local function add(a, b, c) return a + b + c end\nlocal t = {1, 2, 3}\nprint(add(table.unpack(t)))\n",
        "6"
    },
    unpack_empty_slice_nil => {
        "local t = {1, 2, 3}\nlocal a, b = table.unpack(t, 2, 1)\nprint(tostring(a))\n",
        "nil"
    },
    unpack_to_locals => {
        "local a, b, c = table.unpack({7, 8, 9})\nprint(a .. \",\" .. b .. \",\" .. c)\n",
        "7,8,9"
    },
    unpack_discard_extra => {
        "local a, b = table.unpack({1, 2, 3})\nprint(a .. \",\" .. b)\n",
        "1,2"
    },
    unpack_missing_nil => {
        "local a, b, c = table.unpack({1, 2})\nprint(tostring(c))\n",
        "nil"
    },
    unpack_concat_spread => {
        "local function join(...) return table.concat({...}, \"-\") end\nprint(join(table.unpack({\"a\", \"b\", \"c\"})))\n",
        "a-b-c"
    },
}
