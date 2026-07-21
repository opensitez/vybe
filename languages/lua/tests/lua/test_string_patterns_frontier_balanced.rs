//! String frontier pattern `%f[set]` and balanced match `%b` (Lua 5.x §6.4.1)

lua_print! {
    balanced_match_parens => {
        "print(string.match(\"(hello)\", \"%b()\"))\n",
        "(hello)"
    },
    balanced_match_nested => {
        "print(string.match(\"((a)(b))\", \"%b()\"))\n",
        "((a)(b))"
    },
    balanced_match_curly => {
        "print(string.match(\"{one {two}}\", \"%b{}\"))\n",
        ""
    },
    balanced_match_nil => {
        "print(tostring(string.match(\"no parens\", \"%b()\")))\n",
        "nil"
    },
    frontier_pattern_word => {
        "local s = \"THE END\"\nprint(string.find(s, \"%f[%a]%u+\"))\n",
        "1\t3"
    },
    frontier_before_lower => {
        "local s = \"hello world\"\nlocal t = {}\nfor w in string.gmatch(s, \"%f[%a]%a+\") do t[#t+1] = w end\nprint(#t)\n",
        "2"
    },
    frontier_matches_start => {
        "local s = \"abc def\"\nlocal n = 0\nfor _ in string.gmatch(s, \"%f[%a]%a+\") do n = n + 1 end\nprint(n)\n",
        "2"
    },
    balanced_match_square => {
        "print(string.match(\"[inner]\", \"%b[]\"))\n",
        "[inner]"
    },
    balanced_match_span => {
        "local m = string.match(\"(x+y)\", \"%b()\")\nprint(m)\n",
        "(x+y)"
    },
}
