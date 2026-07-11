//! Super table library assertion tests (Lua 5.x §6.6)

lua_print! {
    tbl_sup_insert_1 => {
        "local t = {}\ntable.insert(t, 1)\nprint(t[1])\n",
        "1"
    },
    tbl_sup_insert_2 => {
        "local t = {}\ntable.insert(t, 2)\nprint(t[1])\n",
        "2"
    },
    tbl_sup_insert_3 => {
        "local t = {}\ntable.insert(t, 3)\nprint(t[1])\n",
        "3"
    },
    tbl_sup_insert_4 => {
        "local t = {}\ntable.insert(t, 4)\nprint(t[1])\n",
        "4"
    },
    tbl_sup_insert_5 => {
        "local t = {}\ntable.insert(t, 5)\nprint(t[1])\n",
        "5"
    },
    tbl_sup_insert_6 => {
        "local t = {}\ntable.insert(t, 6)\nprint(t[1])\n",
        "6"
    },
    tbl_sup_insert_7 => {
        "local t = {}\ntable.insert(t, 7)\nprint(t[1])\n",
        "7"
    },
    tbl_sup_insert_8 => {
        "local t = {}\ntable.insert(t, 8)\nprint(t[1])\n",
        "8"
    },
    tbl_sup_insert_9 => {
        "local t = {}\ntable.insert(t, 9)\nprint(t[1])\n",
        "9"
    },
    tbl_sup_remove_1 => {
        "local t = {1}\ntable.remove(t)\nprint(#t)\n",
        "0"
    },
    tbl_sup_remove_2 => {
        "local t = {1, 2}\ntable.remove(t)\nprint(#t)\n",
        "1"
    },
    tbl_sup_remove_3 => {
        "local t = {1, 2, 3}\ntable.remove(t)\nprint(#t)\n",
        "2"
    },
    tbl_sup_remove_4 => {
        "local t = {1, 2, 3, 4}\ntable.remove(t)\nprint(#t)\n",
        "3"
    },
    tbl_sup_remove_5 => {
        "local t = {1, 2, 3, 4, 5}\ntable.remove(t)\nprint(#t)\n",
        "4"
    },
    tbl_sup_remove_6 => {
        "local t = {1, 2, 3, 4, 5, 6}\ntable.remove(t)\nprint(#t)\n",
        "5"
    },
    tbl_sup_remove_7 => {
        "local t = {1, 2, 3, 4, 5, 6, 7}\ntable.remove(t)\nprint(#t)\n",
        "6"
    },
    tbl_sup_remove_8 => {
        "local t = {1, 2, 3, 4, 5, 6, 7, 8}\ntable.remove(t)\nprint(#t)\n",
        "7"
    },
    tbl_sup_remove_9 => {
        "local t = {1, 2, 3, 4, 5, 6, 7, 8, 9}\ntable.remove(t)\nprint(#t)\n",
        "8"
    },
}
