//! Exhaustive type conversion, representation, and lexical checks (Lua 5.x §2.1)

lua_print! {
type_nil_check => { "print(type(nil))\n", "nil" },
type_bool_true => { "print(type(true))\n", "boolean" },
type_bool_false => { "print(type(false))\n", "boolean" },
type_number_int => { "print(type(42))\n", "number" },
type_number_flt => { "print(type(3.14))\n", "number" },
type_string_check => { "print(type(\"abc\"))\n", "string" },
type_table_check => { "print(type({}))\n", "table" },
type_function_check => { "print(type(function() end))\n", "function" },
type_thread_check => { "print(type(coroutine.create(function() end)))\n", "thread" },
num_hex_format => { "print(0xFF)\n", "255" },
num_hex_float => { "print(0x1.8p1)\n", "3.0" },
num_scientific => { "print(1e2)\n", "100.0" },
num_scientific_neg => { "print(1e-2)\n", "0.01" },
str_escape_newline => { "print(\"a\\nb\")\n", "a\nb" },
str_escape_quote => { "print(\"a\\\"b\")\n", "a\"b" },
str_long_brackets => { "print([[a\nb]])\n", "a\nb" },
str_long_brackets_eq => { "print([=[a[[]]b]=])\n", "a[[]]b" },
coercion_str_to_num => { "print(\"10\" + 5)\n", "15.0" },
coercion_num_to_str => { "print(10 .. \"abc\")\n", "10abc" },
table_index_num => { "local t = {}; t[1] = \"a\"; print(t[1])\n", "a" },
table_index_str => { "local t = {}; t[\"x\"] = \"b\"; print(t.x)\n", "b" },
table_index_bool => { "local t = {}; t[true] = \"c\"; print(t[true])\n", "c" },
table_index_tbl => { "local t = {}; local k = {}; t[k] = \"d\"; print(t[k])\n", "d" },
table_len_sequence => { "print(#{10, 20, 30})\n", "3" },
table_len_hash_only => { "print(#({x=1, y=2}))\n", "0" },
fn_closure_nested => {
    "local function f(x)\n  return function(y) return x + y end\nend\nprint(f(5)(10))\n",
    "15"
},
fn_vararg_sum => {
    "local function f(...) return select(\"#\", ...) end\nprint(f(1, 2, 3))\n",
    "3"
} }
