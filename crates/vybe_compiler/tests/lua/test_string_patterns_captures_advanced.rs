//! Pattern matching captures, nested groups, and empty captures (Lua 5.x §6.4.1)

lua_print! {
    pattern_nested_groups => {
        "local outer, inner = string.match(\"hello(world)\", \"(%a+%((%a+)%))\")\nprint(outer .. \",\" .. inner)\n",
        "hello(world),world"
    },
    pattern_empty_capture_pos => {
        "local pos = string.match(\"abc\", \"b()\")\nprint(pos)\n",
        "3"
    },
    pattern_zero_or_more => {
        "local cap = string.match(\"abc\", \"a(b*)c\")\nprint(cap)\n",
        "b"
    },
    pattern_one_or_more => {
        "local cap = string.match(\"abbbc\", \"a(b+)c\")\nprint(cap)\n",
        "bbb"
    },
    pattern_optional_absent => {
        "local cap = string.match(\"ac\", \"a(b?)c\")\nprint(cap)\n",
        ""
    },
    pattern_percent_literal_match => {
        "local cap = string.match(\"a%b\", \"a(%%b)\")\nprint(cap)\n",
        "%b"
    },
    pattern_balanced_empty_tag => {
        "local cap = string.match(\"<>\", \"(%b<>)\")\nprint(cap)\n",
        "<>"
    },
    pattern_frontier_start => {
        "local pos = string.find(\"abc\", \"%f[%a]a\")\nprint(pos)\n",
        "1"
    },
}
