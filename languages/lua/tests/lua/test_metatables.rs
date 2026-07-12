//! Metatables and metamethods — Lua 5.x manual §2.4, §3.3.7, §6.1.

lua_print! {
    setmetatable_roundtrip_via_getmetatable => {
        "local t={}\nlocal m={}\nsetmetatable(t,m)\nprint(getmetatable(t)==m)\n",
        "true"
    },
    __index_table_fallback_reads_missing_key => {
        "local defaults={x=1}\nlocal t=setmetatable({}, {__index=defaults})\nprint(t.x)\n",
        "1"
    },
    __index_function_computes_missing_key => {
        "local t=setmetatable({}, {__index=function(tbl,k) return #k end})\nprint(t.abc)\n",
        "3"
    },
    __add_metamethod_on_tables => {
        "local mt={__add=function(a,b) return a.v+b.v end}\nlocal a=setmetatable({v=2},mt)\nlocal b=setmetatable({v=5},mt)\nprint((a+b).v)\n",
        "7"
    },
    __eq_metamethod_can_make_equal => {
        "local mt={__eq=function(a,b) return a.id==b.id end}\nlocal a=setmetatable({id=1},mt)\nlocal b=setmetatable({id=1},mt)\nprint(a==b)\n",
        "true"
    },
    rawget_bypasses_metamethod => {
        "local t=setmetatable({x=1},{__index={y=9}})\nprint(rawget(t,\"y\"))\n",
        "nil"
    },
    rawset_bypasses_newindex => {
        "local t=setmetatable({},{__newindex=function() error(\"blocked\") end})\nrawset(t,\"k\",1)\nprint(t.k)\n",
        "1"
    },
    __len_metamethod_overrides_length_operator => {
        "local t=setmetatable({},{__len=function() return 5 end})\nprint(#t)\n",
        "5"
    },
    __tostring_metamethod_used_by_tostring => {
        "local t=setmetatable({},{__tostring=function() return \"tbl\" end})\nprint(tostring(t))\n",
        "tbl"
    },
    __newindex_table_stores_fallback_keys => {
        "local store={}\nlocal t=setmetatable({},{__newindex=store})\nt.x=1\nprint(store.x)\n",
        "1"
    },
    __call_metamethod_invokes_on_function_syntax => {
        "local t=setmetatable({}, {__call=function(_,a) return a*2 end})\nprint(t(4))\n",
        "8"
    },
    metatable_nil_means_no_metamethods => {
        "print(tostring(getmetatable({})))\n",
        "nil"
    },
    __lt_metamethod_orders_values => {
        "local mt = {__lt = function(a,b) return a.v < b.v end}\nlocal a = setmetatable({v=1}, mt)\nlocal b = setmetatable({v=2}, mt)\nprint(a < b)\n",
        "true"
    },
    __concat_metamethod_joins_values => {
        "local mt = {__concat = function(a,b) return a.v .. b.v end}\nlocal a = setmetatable({v=\"a\"}, mt)\nlocal b = setmetatable({v=\"b\"}, mt)\nprint(a .. b)\n",
        "ab"
    },
    __metatable_field_hides_metatable => {
        "local hidden = {}\nlocal t = setmetatable({}, {__metatable=hidden})\nprint(getmetatable(t) == hidden)\n",
        "true"
    },
    __index_metamethod_on_metatable_itself => {
        "local mt = {__index = function(_, k) return k .. \"!\" end}\nlocal t = setmetatable({}, mt)\nprint(t.hello)\n",
        "hello!"
    },
    __newindex_routes_writes_to_external_store => {
        "local data = {}\nlocal t = setmetatable({}, {__newindex = data})\nt.a = 1\nprint(data.a)\n",
        "1"
    },
    __tostring_for_custom_print_representation => {
        "local t = setmetatable({}, {__tostring = function() return \"<obj>\" end})\nprint(tostring(t))\n",
        "<obj>"
    },
    __add_with_two_operands_of_same_metatable => {
        "local mt = {__add = function(a, b) return a.n + b.n end}\nlocal x = setmetatable({n = 2}, mt)\nlocal y = setmetatable({n = 3}, mt)\nprint(x + y)\n",
        "5"
    },
    __eq_requires_same_metatable_for_custom_equality => {
        "local mt = {__eq = function(a, b) return a.id == b.id end}\nprint(setmetatable({id=1}, mt) == setmetatable({id=1}, mt))\n",
        "true"
    },
    __len_customizes_length_operator => {
        "local t = setmetatable({}, {__len = function() return 42 end})\nprint(#t)\n",
        "42"
    },
    __call_allows_table_as_function => {
        "local t = setmetatable({}, {__call = function(_, x) return x + 1 end})\nprint(t(4))\n",
        "5"
    },
    rawget_reads_own_field_not_index_fallback => {
        "local t = setmetatable({x = 1}, {__index = {x = 9}})\nprint(rawget(t, \"x\"))\n",
        "1"
    },
    rawset_writes_without_triggering_newindex => {
        "local log = 0\nlocal t = setmetatable({}, {__newindex = function() log = log + 1 end})\nrawset(t, \"k\", 1)\nprint(log)\n",
        "0"
    },
    setmetatable_returns_original_table => {
        "local t = {}\nlocal r = setmetatable(t, {})\nprint(r == t)\n",
        "true"
    },
    metatable_protected_by_metatable_field => {
        "local mt = {__metatable = \"locked\"}\nlocal ok = pcall(function() setmetatable({}, mt) end)\nprint(ok)\n",
        "false"
    },
    metatable_len_metamethod => {
        "local t = setmetatable({}, {__len = function() return 42 end})\nprint(#t)\n",
        "42"
    },
    metatable_concat_metamethod => {
        "local mt = {__concat = function(a, b) return a.v .. b.v end}\nlocal a = setmetatable({v='hello'}, mt)\nlocal b = setmetatable({v=' world'}, mt)\nprint(a .. b)\n",
        "hello world"
    },
    metatable_index_chain_three_levels => {
        "local base = {method = function() return 'base' end}\nlocal mid = setmetatable({}, {__index = base})\nlocal top = setmetatable({}, {__index = mid})\nprint(top.method())\n",
        "base"
    },
    metatable_newindex_intercepts_write => {
        "local log = {}\nlocal t = setmetatable({}, {\n  __newindex = function(tbl, key, val)\n    log[#log+1] = key .. '=' .. tostring(val)\n    rawset(tbl, key, val)\n  end\n})\nt.x = 10\nt.y = 20\nprint(table.concat(log, ','))\n",
        "x=10,y=20"
    },
    rawset_does_not_invoke_newindex_when_writing => {
        "local blocked = false\nlocal t = setmetatable({}, {\n  __newindex = function() blocked = true end\n})\nrawset(t, 'key', 'val')\nprint(tostring(blocked) .. ',' .. t.key)\n",
        "false,val"
    },
    rawget_reads_own_field_skipping_index_metamethod => {
        "local fallback_used = false\nlocal t = setmetatable({}, {\n  __index = function() fallback_used = true; return 99 end\n})\nrawset(t, 'x', 42)\nlocal v = rawget(t, 'x')\nprint(tostring(fallback_used) .. ',' .. v)\n",
        "false,42"
    },
    metatable_call_metamethod_on_table => {
        "local callable = setmetatable({}, {\n  __call = function(self, a, b) return a + b end\n})\nprint(callable(3, 4))\n",
        "7"
    },
    getmetatable_returns_nil_on_plain_table => {
        "local t = {}\nprint(tostring(getmetatable(t)))\n",
        "nil"
    },
}
