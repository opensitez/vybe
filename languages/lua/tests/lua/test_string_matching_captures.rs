//! `string.match` — capture groups, greedy/lazy specifiers, and offsets (Lua 5.x §6.4.1)

lua_print! {
    match_two => {
        "local k, v = string.match(\"color=blue\", \"(%a+)=(%a+)\")\nprint(k .. \"=\" .. v)\n",
        "color=blue"
    },
    match_optional => {
        "local m = string.match(\"abc123\", \"%a+(%d*)\")\nprint(m)\n",
        "123"
    },
    match_offset => {
        "print(string.match(\"aabba\", \"b+\", 3))\n",
        "bb"
    },
    match_date => {
        "local y,m,d = string.match(\"2024-07-11\", \"(%d%d%d%d)-(%d%d)-(%d%d)\")\nprint(y .. \",\" .. m .. \",\" .. d)\n",
        "2024,07,11"
    },
    match_no_match_nil => {
        "print(tostring(string.match(\"hello\", \"%d+\")))\n",
        "nil"
    },
    match_lazy_star => {
        "print(string.match(\"<a><b>\", \"<(.-)>\"))\n",
        "a"
    },
    match_greedy_plus => {
        "print(string.match(\"<a><b>\", \"<(.+)>\"))\n",
        "a><b"
    },
    match_anchored => {
        "print(tostring(string.match(\"hello\", \"^hello$\")))\n",
        "hello"
    },
    match_float_extract => {
        "local n = string.match(\"price: 42.5\", \"(%d+%.?%d*)\")\nprint(n)\n",
        "42.5"
    },
    match_iterative_offsets => {
        "local s = \"one two three\"\nlocal words = {}\nfor w in string.gmatch(s, \"%a+\") do\n  words[#words+1] = w\nend\nprint(table.concat(words, \"-\"))\n",
        "one-two-three"
    } }
