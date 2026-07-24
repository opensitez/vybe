//! Table constructors — list, record, mixed forms (Lua 5.x manual §3.4.8).

lua_print! {
    constructor_list_part_preserves_order => {
        "local t = {3, 1, 4}\nprint(t[1]..\",\"..t[2]..\",\"..t[3])\n",
        "3,1,4"
    },
    constructor_record_part_field_access => {
        "local t = {name=\"lua\", ver=5}\nprint(t.name..t.ver)\n",
        "lua5"
    },
    constructor_semicolon_separates_parts => {
        "local t = {1, 2; a=3}\nprint(t[2]+t.a)\n",
        "5"
    },
    constructor_trailing_comma_allowed => {
        "local t = {1, 2,}\nprint(#t)\n",
        "2"
    },
    constructor_bracket_key_expression => {
        "local k=\"x\"\nlocal t = {[k]=9}\nprint(t.x)\n",
        "9"
    },
    constructor_numeric_key_in_brackets => {
        "local t = {[2]=\"b\", [1]=\"a\"}\nprint(t[1]..t[2])\n",
        "ab"
    },
    constructor_function_field => {
        "local t = {run=function() return 4 end}\nprint(t.run())\n",
        "4"
    },
    constructor_nested_table_field => {
        "local t = {inner={v=2}}\nprint(t.inner.v)\n",
        "2"
    },
    constructor_list_then_record => {
        "local t = {10, x=20}\nprint(t[1]+t.x)\n",
        "30"
    },
    constructor_empty_yields_empty_table => {
        "print(next({})==nil)\n",
        "true"
    },
    constructor_duplicate_keys_last_wins => {
        "local t = {a=1, a=2}\nprint(t.a)\n",
        "2"
    },
    constructor_ellipsis_not_in_table_literal => {
        "local parts = {1,2,3}\nlocal t = {table.unpack(parts)}\nprint(t[3])\n",
        "3"
    },
    constructor_boolean_keys_via_brackets => {
        "local t = {[true]=\"yes\", [false]=\"no\"}\nprint(t[true])\n",
        "yes"
    },
    constructor_nil_key_entry_errors => {
        "local ok = pcall(function() return {[nil]=1, a=2} end)\nprint(tostring(ok))\n",
        "false"
    },
    constructor_string_key_without_brackets => {
        "local t = {hello=\"world\"}\nprint(t.hello)\n",
        "world"
    },
}
