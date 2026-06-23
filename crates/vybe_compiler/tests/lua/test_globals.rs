//! Global environment — undeclared names, `_G` (Lua 5.x manual §2.2).

lua_print! {
    undeclared_assignment_creates_global => {
        "foo = 10\nprint(foo)\n",
        "10"
    },
    global_persists_across_statements => {
        "bar = 1\nbar = bar + 2\nprint(bar)\n",
        "3"
    },
    read_global_before_local_shadows_later => {
        "baz = 5\nlocal baz = 1\nprint(baz)\n",
        "1"
    },
    _g_table_holds_globals => {
        "_G.answer = 42\nprint(_G.answer)\n",
        "42"
    },
    rawget_on_global_table_reads_value => {
        "xyzzy = 7\nprint(rawget(_G, \"xyzzy\"))\n",
        "7"
    },
    nil_global_read_returns_nil => {
        "print(tostring(no_such_global))\n",
        "nil"
    },
    assign_to_global_without_local_declaration => {
        "count = 1\ncount = count + 1\nprint(count)\n",
        "2"
    },
    read_global_in_function_body => {
        "factor = 2\nfunction scale(x) return x * factor end\nprint(scale(3))\n",
        "6"
    },
    global_string_and_number_mixed => {
        "name = \"lua\"\nprint(name .. 5)\n",
        "lua5"
    },
}
