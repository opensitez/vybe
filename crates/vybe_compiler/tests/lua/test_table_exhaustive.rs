//! Table library exhaustive tests: insert, remove, concat, sort, pack, unpack, move (Lua 5.x §6.6)

lua_print! {
    tbl_exh_insert_tail => {
        "local t = {10, 20}\ntable.insert(t, 30)\nprint(t[1], t[2], t[3])\n",
        "10\t20\t30"
    },
    tbl_exh_insert_pos => {
        "local t = {10, 20}\ntable.insert(t, 2, 15)\nprint(t[1], t[2], t[3])\n",
        "10\t15\t20"
    },
    tbl_exh_remove_tail => {
        "local t = {10, 20, 30}\nlocal v = table.remove(t)\nprint(v, #t)\n",
        "30\t2"
    },
    tbl_exh_remove_pos => {
        "local t = {10, 20, 30}\nlocal v = table.remove(t, 2)\nprint(v, t[1], t[2])\n",
        "20\t10\t30"
    },
    tbl_exh_concat_default => {
        "print(table.concat({\"a\", \"b\", \"c\"}))\n",
        "abc"
    },
    tbl_exh_concat_sep => {
        "print(table.concat({\"a\", \"b\", \"c\"}, \"-\"))\n",
        "a-b-c"
    },
    tbl_exh_concat_range => {
        "print(table.concat({\"a\", \"b\", \"c\", \"d\"}, \"-\", 2, 3))\n",
        "b-c"
    },
    tbl_exh_sort_default => {
        "local t = {3, 1, 2}\ntable.sort(t)\nprint(t[1], t[2], t[3])\n",
        "1\t2\t3"
    },
    tbl_exh_sort_comparator => {
        "local t = {1, 3, 2}\ntable.sort(t, function(a, b) return a > b end)\nprint(t[1], t[2], t[3])\n",
        "3\t2\t1"
    },
    tbl_exh_pack_count => {
        "local t = table.pack(10, nil, 30)\nprint(t.n, t[1], t[2], t[3])\n",
        "3\t10\tnil\t30"
    },
    tbl_exh_unpack_defaults => {
        "local a, b = table.unpack({10, 20})\nprint(a, b)\n",
        "10\t20"
    },
    tbl_exh_unpack_range => {
        "local a, b = table.unpack({10, 20, 30, 40}, 2, 3)\nprint(a, b)\n",
        "20\t30"
    },
    tbl_exh_move_basic => {
        "local a = {10, 20, 30}\nlocal b = {}\ntable.move(a, 1, 3, 1, b)\nprint(b[1], b[2], b[3])\n",
        "10\t20\t30"
    },
}
