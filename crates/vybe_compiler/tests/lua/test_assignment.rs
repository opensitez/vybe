//! Assignment — multiple values, evaluation order (Lua 5.x manual §3.3.3).

lua_print! {
    multiple_assignment_swaps_variables => {
        "local a, b = 1, 2\na, b = b, a\nprint(a .. \",\" .. b)\n",
        "2,1"
    },
    extra_rhs_values_are_discarded => {
        "local a, b = 1, 2, 3\nprint(a .. \",\" .. b)\n",
        "1,2"
    },
    missing_rhs_values_become_nil => {
        "local a, b, c = 1\nprint(tostring(c))\n",
        "nil"
    },
    function_return_fills_multiple_locals => {
        "local function pair() return 10, 20 end\nlocal x, y = pair()\nprint(x + y)\n",
        "30"
    },
    indexed_assignment_to_table_field => {
        "local t = {}\nt[\"k\"] = 7\nprint(t.k)\n",
        "7"
    },
    compound_table_field_update => {
        "local t = {n = 1}\nt.n = t.n + 2\nprint(t.n)\n",
        "3"
    },
    local_list_declares_without_initializers => {
        "local a, b\nc = 1\nb = 2\na = b + c\nprint(a)\n",
        "3"
    },
    assignment_evaluates_rhs_before_lhs => {
        "local t = {1, 2, 3}\nlocal i = 1\nt[i], i = t[i + 1], i + 1\nprint(t[1] .. \",\" .. i)\n",
        "2,2"
    },
    upvalue_assignment_through_closure => {
        "local n = 0\nlocal function set(v) n = v end\nset(4)\nprint(n)\n",
        "4"
    },
    global_assignment_from_function => {
        "function setg() gmark = 11 end\nsetg()\nprint(gmark)\n",
        "11"
    },
    destructuring_table_fields_into_locals => {
        "local t = {x = 1, y = 2}\nlocal a, b = t.x, t.y\nprint(a + b)\n",
        "3"
    },
    update_table_field_with_self_reference => {
        "local cfg = {count = 1}\ncfg.count = cfg.count + 1\nprint(cfg.count)\n",
        "2"
    },
    assign_from_function_multiple_returns_to_locals => {
        "local function minmax(a, b)\n  if a < b then return a, b else return b, a end\nend\nlocal lo, hi = minmax(3, 1)\nprint(lo .. \",\" .. hi)\n",
        "1,3"
    },
}
