//! Table library extended tests — sort, concat, insert, remove, move, unpack, pack (Lua 5.x §6.6)

lua_print! {
    table_concat_delim => { "print(table.concat({\"a\", \"b\", \"c\"}, \"-\"))\n", "a-b-c" },
    table_concat_subset => { "print(table.concat({\"a\", \"b\", \"c\", \"d\"}, \"-\", 2, 3))\n", "b-c" },
    table_concat_empty => { "print(table.concat({}))\n", "" },
    table_concat_non_string => { "print(table.concat({1, 2, 3}, \",\"))\n", "1,2,3" },
    table_insert_tail => {
        "local t = {10, 20}\ntable.insert(t, 30)\nprint(t[3])\n",
        "30"
    },
    table_insert_pos => {
        "local t = {10, 20}\ntable.insert(t, 2, 15)\nprint(t[1] .. \",\" .. t[2] .. \",\" .. t[3])\n",
        "10,15,20"
    },
    table_remove_tail => {
        "local t = {10, 20, 30}\nlocal r = table.remove(t)\nprint(r .. \",\" .. #t)\n",
        "30,2"
    },
    table_remove_pos => {
        "local t = {10, 20, 30}\nlocal r = table.remove(t, 2)\nprint(r .. \",\" .. t[2])\n",
        "20,30"
    },
    table_move_elements => {
        "local a = {10, 20, 30}\nlocal b = {}\ntable.move(a, 1, 3, 1, b)\nprint(b[1] .. \",\" .. b[3])\n",
        "10,30"
    },
    table_move_overwrite => {
        "local t = {10, 20, 30}\ntable.move(t, 1, 2, 2)\nprint(t[1] .. \",\" .. t[2] .. \",\" .. t[3])\n",
        "10,10,20"
    },
    table_sort_default => {
        "local t = {3, 1, 2}\ntable.sort(t)\nprint(t[1] .. \",\" .. t[3])\n",
        "1,3"
    },
    table_sort_comparator => {
        "local t = {1, 3, 2}\ntable.sort(t, function(a, b) return a > b end)\nprint(t[1] .. \",\" .. t[3])\n",
        "3,1"
    },
    table_pack_count => {
        "local t = table.pack(10, nil, 30)\nprint(t.n .. \",\" .. tostring(t[2]))\n",
        "3,nil"
    },
    table_unpack_defaults => {
        "local a, b = table.unpack({10, 20})\nprint(a .. \",\" .. b)\n",
        "10,20"
    },
    table_unpack_indices => {
        "local a, b = table.unpack({10, 20, 30, 40}, 2, 3)\nprint(a .. \",\" .. b)\n",
        "20,30"
    },
}
