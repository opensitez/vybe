//! `string.gmatch` — iterator-based pattern matching (Lua 5.x §6.4.1)

lua_print! {
    gmatch_words => {
        "local t={}\nfor w in string.gmatch(\"one two three\", \"%a+\") do t[#t+1]=w end\nprint(table.concat(t, \",\"))\n",
        "one,two,three"
    },
    gmatch_key_val => {
        "local t={}\nfor k,v in string.gmatch(\"a=1,b=2\", \"(%a+)=(%d+)\") do t[k]=v end\nprint(t[\"a\"] .. \",\" .. t[\"b\"])\n",
        "1,2"
    },
    gmatch_count => {
        "local n=0\nfor _ in string.gmatch(\"1,2,3,4\", \"%d+\") do n=n+1 end\nprint(n)\n",
        "4"
    },
    gmatch_no_match => {
        "local n=0\nfor _ in string.gmatch(\"abc\", \"%d+\") do n=n+1 end\nprint(n)\n",
        "0"
    },
    gmatch_chars => {
        "local r=\"\"\nfor c in string.gmatch(\"lua\", \".\") do r=r..c..\"-\" end\nprint(r)\n",
        "l-u-a-"
    },
    gmatch_lines => {
        "local n=0\nfor _ in string.gmatch(\"a\\nb\\nc\", \"[^\\n]+\") do n=n+1 end\nprint(n)\n",
        "3"
    },
    gmatch_digits => {
        "local s=\"\"\nfor d in string.gmatch(\"abc123def456\", \"%d+\") do s=s..d..\",\" end\nprint(s)\n",
        "123,456,"
    },
    gmatch_non_spaces => {
        "local n=0\nfor _ in string.gmatch(\"  hello  world  \", \"%S+\") do n=n+1 end\nprint(n)\n",
        "2"
    },
    gmatch_no_captures_returns_full_match => {
        "local r={}\nfor m in string.gmatch(\"cat bat sat\", \"%a+at\") do r[#r+1]=m end\nprint(table.concat(r,\",\"))\n",
        "cat,bat,sat"
    },
    gmatch_multiple_captures => {
        "local r={}\nfor a,b in string.gmatch(\"x:1 y:2\", \"(%a):(%d)\") do r[#r+1]=a..b end\nprint(table.concat(r,\",\"))\n",
        "x1,y2"
    },
    gmatch_empty_pattern => {
        "local n=0\nfor _ in string.gmatch(\"ab\", \"\") do n=n+1 end\nprint(n)\n",
        "3"
    } }
