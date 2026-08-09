//! Table constructor syntax: mixed, record, sequence, nested (Lua 5.x §3.4.9)

lua_print! {
constructor_seq_keys => {
    "local t = {\"a\", \"b\", \"c\"}\nprint(t[1] .. t[3])\n",
    "ac"
},
constructor_record => {
    "local t = {x=1, y=2}\nprint(t.x + t.y)\n",
    "3"
},
constructor_explicit => {
    "local t = {[10]=\"ten\", [20]=\"twenty\"}\nprint(t[10])\n",
    "ten"
},
constructor_mixed => {
    "local t = {\"a\", x=99, \"b\"}\nprint(t[1] .. t[2] .. t.x)\n",
    "ab99"
},
constructor_bracket_str => {
    "local t = {[\"hello world\"]=\"hi\"}\nprint(t[\"hello world\"])\n",
    "hi"
},
constructor_nested => {
    "local t = {inner={a=1, b=2}}\nprint(t.inner.a + t.inner.b)\n",
    "3"
},
constructor_fn_val => {
    "local t = {fn=function(x) return x*2 end}\nprint(t.fn(5))\n",
    "10"
},
constructor_trailing_comma => {
    "local t = {1, 2, 3 }\nprint(#t)\n",
    "3"
},
constructor_seq_fn => {
    "local function pair() return 10, 20 end\nlocal t = {pair()}\nprint(#t .. \",\" .. t[2])\n",
    "2,20"
},
constructor_seq_fn_mid => {
    "local function pair() return 10, 20 end\nlocal t = {pair(), 30}\nprint(#t .. \",\" .. t[2])\n",
    "2,30"
},
constructor_int_str_keys => {
    "local t = {[1]=\"idx\", one=\"str\"}\nprint(t[1] .. \",\" .. t.one)\n",
    "idx,str"
} }
