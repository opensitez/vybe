//! Metamethods — arithmetic, comparison, indexing (Lua 5.x manual §2.4, §3.3.7).

lua_print! {
metamethod_sub_on_tables => {
    "local mt={__sub=function(a,b) return a.n-b.n end}\nlocal a=setmetatable({n=5},mt)\nlocal b=setmetatable({n=2},mt)\nprint((a-b).n)\n",
    "3"
},
metamethod_mul_on_tables => {
    "local mt={__mul=function(a,b) return a.n*b.n end}\nprint((setmetatable({n=3},mt)*setmetatable({n=4},mt)).n)\n",
    "12"
},
metamethod_div_on_tables => {
    "local mt={__div=function(a,b) return a.n/b.n end}\nprint((setmetatable({n=8},mt)/setmetatable({n=2},mt)).n)\n",
    "4"
},
metamethod_mod_on_tables => {
    "local mt={__mod=function(a,b) return a.n%b.n end}\nprint((setmetatable({n=10},mt)%setmetatable({n=3},mt)).n)\n",
    "1"
},
metamethod_pow_on_tables => {
    "local mt={__pow=function(a,b) return a.n^b.n end}\nprint((setmetatable({n=2},mt)^setmetatable({n=3},mt)).n)\n",
    "8"
},
metamethod_unm_negates => {
    "local mt={__unm=function(a) return {n=-a.n} end}\nprint((-setmetatable({n=4},mt)).n)\n",
    "-4"
},
metamethod_idiv_floor_divides => {
    "local mt={__idiv=function(a,b) return {n=a.n//b.n} end}\nprint((setmetatable({n=7},mt)//setmetatable({n=2},mt)).n)\n",
    "3"
},
metamethod_band_on_tables => {
    "local mt={__band=function(a,b) return {n=a.n&b.n} end}\nprint((setmetatable({n=6},mt)&setmetatable({n=3},mt)).n)\n",
    "2"
},
metamethod_bor_on_tables => {
    "local mt={__bor=function(a,b) return {n=a.n|b.n} end}\nprint((setmetatable({n=1},mt)|setmetatable({n=2},mt)).n)\n",
    "3"
},
metamethod_bxor_on_tables => {
    "local mt={__bxor=function(a,b) return {n=a.n~b.n} end}\nprint((setmetatable({n=5},mt)~setmetatable({n=3},mt)).n)\n",
    "6"
},
metamethod_bnot_on_tables => {
    "local mt={__bnot=function(a) return {n=~a.n} end}\nprint((~setmetatable({n=0},mt)).n)\n",
    "-1"
},
metamethod_shl_shifts_left => {
    "local mt={__shl=function(a,b) return {n=a.n<<b.n} end}\nprint((setmetatable({n=1},mt)<<setmetatable({n=3},mt)).n)\n",
    "8"
},
metamethod_shr_shifts_right => {
    "local mt={__shr=function(a,b) return {n=a.n>>b.n} end}\nprint((setmetatable({n=8},mt)>>setmetatable({n=1},mt)).n)\n",
    "4"
},
metamethod_concat_joins_payloads => {
    "local mt={__concat=function(a,b) return a.s..b.s end}\nprint(setmetatable({s=\"a\"},mt)..setmetatable({s=\"b\"},mt))\n",
    "ab"
},
metamethod_len_returns_custom => {
    "local mt={__len=function() return 99 end}\nprint(#setmetatable({},mt))\n",
    "99"
},
metamethod_le_orders_values => {
    "local mt={__le=function(a,b) return a.n<=b.n end}\nprint(setmetatable({n=2},mt)<=setmetatable({n=3},mt))\n",
    "true"
},
metamethod_newindex_table_stores_externally => {
    "local store={}\nlocal t=setmetatable({}, {__newindex=store})\nt.k=\"v\"\nprint(store.k)\n",
    "v"
},
metamethod_newindex_function_receives_key_value => {
    "local log=\"\"\nlocal t=setmetatable({}, {__newindex=function(_,k,v) log=k..\"=\"..v end})\nt.x=1\nprint(log)\n",
    "x=1"
},
metamethod_eq_not_used_for_different_types => {
    "print(setmetatable({},{}) == 1)\n",
    "false"
},
metamethod_index_chain_follows_metatable => {
    "local base={x=1}\nlocal mid=setmetatable({}, {__index=base})\nlocal top=setmetatable({}, {__index=mid})\nprint(top.x)\n",
    "1"
},
metamethod_unm_negates_custom_number => {
    "local mt = {__unm = function(a) return -a.n end}\nprint((-setmetatable({n = 5}, mt)).n)\n",
    "-5"
},
metamethod_concat_joins_custom_values => {
    "local mt = {__concat = function(a, b) return a.s .. b.s end}\nlocal a = setmetatable({s = \"x\"}, mt)\nlocal b = setmetatable({s = \"y\"}, mt)\nprint(a .. b)\n",
    "xy"
},
metamethod_lt_orders_custom_objects => {
    "local mt = {__lt = function(a, b) return a.score < b.score end}\nprint(setmetatable({score=1}, mt) < setmetatable({score=2}, mt))\n",
    "true"
},
metamethod_le_includes_equality_case => {
    "local mt = {__le = function(a, b) return a.n <= b.n end}\nprint(setmetatable({n=2}, mt) <= setmetatable({n=2}, mt))\n",
    "true"
},
metamethod_shl_left_shift_on_wrapped_int => {
    "local mt = {__shl = function(a, b) return {n = a.n << b.n} end}\nprint((setmetatable({n=1}, mt) << setmetatable({n=2}, mt)).n)\n",
    "4"
},
metamethod_shr_right_shift_on_wrapped_int => {
    "local mt = {__shr = function(a, b) return {n = a.n >> b.n} end}\nprint((setmetatable({n=8}, mt) >> setmetatable({n=1}, mt)).n)\n",
    "4"
},
metamethod_band_bitwise_and => {
    "local mt = {__band = function(a, b) return {n = a.n & b.n} end}\nprint((setmetatable({n=0xF}, mt) & setmetatable({n=0x3}, mt)).n)\n",
    "3"
},
metamethod_bor_bitwise_or => {
    "local mt = {__bor = function(a, b) return {n = a.n | b.n} end}\nprint((setmetatable({n=1}, mt) | setmetatable({n=2}, mt)).n)\n",
    "3"
},
metamethod_bxor_bitwise_xor => {
    "local mt = {__bxor = function(a, b) return {n = a.n ~ b.n} end}\nprint((setmetatable({n=5}, mt) ~ setmetatable({n=3}, mt)).n)\n",
    "6"
},
metamethod_idiv_floor_division => {
    "local mt = {__idiv = function(a, b) return {n = a.n // b.n} end}\nprint((setmetatable({n=7}, mt) // setmetatable({n=2}, mt)).n)\n",
    "3"
},
metamethod_index_function_receives_table_and_key => {
    "local t = setmetatable({}, {__index = function(tbl, k) return \"k:\" .. k end})\nprint(t.foo)\n",
    "k:foo"
} }
