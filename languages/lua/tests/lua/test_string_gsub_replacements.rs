//! `string.gsub` with function replacement and count limit (Lua 5.x §6.4)

lua_print! {
    gsub_fn => {
        "local r = string.gsub(\"hello world\", \"%a+\", function(w) return w:upper() end)\nprint(r)\n",
        "HELLO WORLD"
    },
    gsub_fn_captures => {
        "local r = string.gsub(\"a=1 b=2\", \"(%a+)=(%d+)\", function(k, v) return k..\"[\"..v..\"]\" end)\nprint(r)\n",
        "a[1] b[2]"
    },
    gsub_fn_nil => {
        "local r = string.gsub(\"abc\", \"%a\", function(m) if m == \"b\" then return nil end return m:upper() end)\nprint(r)\n",
        "AbC"
    },
    gsub_fn_false => {
        "local r = string.gsub(\"abc\", \"%a\", function(m) if m == \"b\" then return false end return m:upper() end)\nprint(r)\n",
        "AbC"
    },
    gsub_tbl => {
        "local t = {cat=\"CAT\", dog=\"DOG\"}\nlocal r = string.gsub(\"my cat and dog\", \"%a+\", t)\nprint(r)\n",
        "my CAT and DOG"
    },
    gsub_tbl_missing => {
        "local t = {hello=\"HI\"}\nlocal r = string.gsub(\"hello world\", \"%a+\", t)\nprint(r)\n",
        "HI world"
    },
    gsub_limit => {
        "local r = string.gsub(\"aaa\", \"a\", \"b\", 2)\nprint(r)\n",
        "bba"
    },
    gsub_count_ret => {
        "local _, n = string.gsub(\"banana\", \"a\", \"\")\nprint(n)\n",
        "3"
    },
    gsub_empty_pat => {
        "local r = string.gsub(\"ab\", \"\", \"-\")\nprint(r)\n",
        "-a-b-"
    },
    gsub_backref => {
        "local r = string.gsub(\"2024-07-11\", \"(%d+)-(%d+)-(%d+)\", \"%3/%2/%1\")\nprint(r)\n",
        "11/07/2024"
    },
    gsub_whole_match => {
        "local r = string.gsub(\"cat\", \"%a+\", \"[%0]\")\nprint(r)\n",
        "[cat]"
    },
}
