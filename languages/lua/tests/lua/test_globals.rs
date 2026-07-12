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
    setting_global_via_g_table_is_readable_by_name => {
        "_G.my_global_val = 77\nprint(my_global_val)\n",
        "77"
    },
    deleting_global_by_assigning_nil => {
        "some_global = 'hello'\nsome_global = nil\nprint(tostring(some_global))\n",
        "nil"
    },
    global_set_inside_pcall_persists_after_success => {
        "pcall(function() side_effect_global = 'set' end)\nprint(side_effect_global)\n",
        "set"
    },
    version_global_is_a_string => {
        "print(type(_VERSION))\n",
        "string"
    },
    g_is_same_reference_as_global_env => {
        "print(_G == _G)\n",
        "true"
    },
    rawset_on_g_creates_accessible_global => {
        "rawset(_G, 'injected', 123)\nprint(injected)\n",
        "123"
    },
    global_read_from_deeply_nested_function => {
        "deeply_nested = 'found'\nlocal function a()\n  local function b()\n    local function c() return deeply_nested end\n    return c()\n  end\n  return b()\nend\nprint(a())\n",
        "found"
    },
    global_update_visible_to_subsequent_function_call => {
        "shared = 0\nlocal function inc() shared = shared + 1 end\nlocal function get() return shared end\ninc(); inc(); inc()\nprint(get())\n",
        "3"
    },
}
