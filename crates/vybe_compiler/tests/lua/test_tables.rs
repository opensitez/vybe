//! Tables — constructors, indexing, length, table library (Lua 5.x manual §3.4.8–3.4.9, §6.1).

lua_print! {
    table_constructor_empty_length => {
        "local t = {}\nprint(#t)\n",
        "0"
    },
    table_array_part_one_based_index => {
        "local t = {}\nt[1] = \"a\"\nt[2] = \"b\"\nprint(t[1] .. t[2])\n",
        "ab"
    },
    table_field_syntax_constructor => {
        "local t = {x = 1, y = 2}\nprint(t.x + t.y)\n",
        "3"
    },
    table_list_and_record_constructor_mix => {
        "local t = {10, a = 5}\nprint(t[1] + t.a)\n",
        "15"
    },
    table_insert_appends_element => {
        "local t = {}\ntable.insert(t, \"x\")\ntable.insert(t, \"y\")\nprint(t[1] .. t[2])\n",
        "xy"
    },
    table_insert_at_interior_position => {
        "local t = {1, 3}\ntable.insert(t, 2, 2)\nprint(table.concat(t, \",\"))\n",
        "1,2,3"
    },
    table_concat_joins_array_part => {
        "local t = {}\nt[1] = \"a\"\nt[2] = \"b\"\nprint(table.concat(t, \",\"))\n",
        "a,b"
    },
    table_concat_with_start_end_range => {
        "local t = {\"a\", \"b\", \"c\"}\nprint(table.concat(t, \"\", 2, 3))\n",
        "bc"
    },
    table_remove_from_end => {
        "local t = {}\nt[1] = 10\nt[2] = 20\nprint(table.remove(t))\n",
        "20"
    },
    table_remove_at_index_shifts_elements => {
        "local t = {10, 20, 30}\ntable.remove(t, 2)\nprint(t[2])\n",
        "30"
    },
    table_sort_orders_numbers_ascending => {
        "local t = {3, 1, 2}\ntable.sort(t)\nprint(table.concat(t, \",\"))\n",
        "1,2,3"
    },
    length_operator_on_sequential_table => {
        "local t = {}\nt[1] = 1\nt[2] = 2\nt[3] = 3\nprint(#t)\n",
        "3"
    },
    length_operator_stops_at_first_nil_hole => {
        "local t = {}\nt[1] = 1\nt[2] = nil\nt[3] = 3\nprint(#t)\n",
        "1"
    },
    missing_key_read_yields_nil => {
        "local t = {}\nprint(tostring(t.missing))\n",
        "nil"
    },
    table_pack_sets_n_field => {
        "local p = table.pack(1, 2, 3)\nprint(p.n)\n",
        "3"
    },
    table_unpack_returns_multiple_values => {
        "local a, b = table.unpack({9, 8})\nprint(a + b)\n",
        "17"
    },
    table_move_copies_slice_to_destination => {
        "local a = {1, 2, 3, 4}\ntable.move(a, 2, 3, 1, a)\nprint(table.concat(a, \",\"))\n",
        "2,3,3,4"
    },
    table_sort_with_custom_comparator_descending => {
        "local t = {3, 1, 2}\ntable.sort(t, function(a,b) return a > b end)\nprint(table.concat(t, \",\"))\n",
        "3,2,1"
    },
    nested_table_field_access => {
        "local t = { inner = { value = 8 } }\nprint(t.inner.value)\n",
        "8"
    },
    bracket_key_with_expression => {
        "local k = \"x\"\nlocal t = {}\nt[k] = 7\nprint(t.x)\n",
        "7"
    },
    table_stores_function_value => {
        "local t = { f = function(x) return x + 1 end }\nprint(t.f(4))\n",
        "5"
    },
    array_index_zero_is_not_sequence_part => {
        "local t = {}\nt[0] = \"z\"\nt[1] = \"a\"\nprint(t[1])\n",
        "a"
    },
    table_unpack_with_index_range => {
        "local a, b = table.unpack({10, 20, 30}, 2, 3)\nprint(a + b)\n",
        "50"
    },
    table_concat_on_empty_array_is_empty_string => {
        "print(table.concat({}))\n",
        ""
    },
    table_sort_orders_strings_lexicographically => {
        "local t = {\"b\", \"a\", \"c\"}\ntable.sort(t)\nprint(table.concat(t, \",\"))\n",
        "a,b,c"
    },
    table_insert_omitting_position_appends => {
        "local t = {1}\ntable.insert(t, 2)\nprint(t[2])\n",
        "2"
    },
    assigning_to_missing_key_creates_entry => {
        "local t = {}\nt.newkey = 4\nprint(t.newkey)\n",
        "4"
    },
    hash_part_does_not_affect_length_operator => {
        "local t = {1, 2}\nt.hidden = 99\nprint(#t)\n",
        "2"
    },
    dot_and_bracket_access_same_field => {
        "local t = {name = \"lua\"}\nprint(t.name == t[\"name\"])\n",
        "true"
    },
    append_by_assigning_next_index => {
        "local t = {1, 2}\nt[#t + 1] = 3\nprint(table.concat(t, \",\"))\n",
        "1,2,3"
    },
    table_as_map_counts_entries_with_pairs => {
        "local counts = {a = 1, b = 2}\nlocal n = 0\nfor _ in pairs(counts) do n = n + 1 end\nprint(n)\n",
        "2"
    },
    build_lookup_table_from_parallel_arrays => {
        "local keys = {\"a\", \"b\"}\nlocal vals = {1, 2}\nlocal map = {}\nfor i = 1, #keys do map[keys[i]] = vals[i] end\nprint(map.b)\n",
        "2"
    },
    use_table_as_stack_with_push_pop => {
        "local st = {}\ntable.insert(st, 10)\ntable.insert(st, 20)\nprint(table.remove(st))\n",
        "20"
    },
    use_table_as_queue_with_remove_first => {
        "local q = {1, 2, 3}\nprint(table.remove(q, 1))\n",
        "1"
    },
    update_nested_field_on_record => {
        "local user = {profile = {name = \"ada\"}}\nuser.profile.name = \"bob\"\nprint(user.profile.name)\n",
        "bob"
    },
    clear_array_by_setting_length_hack_via_remove => {
        "local t = {1, 2, 3}\nwhile #t > 0 do table.remove(t) end\nprint(#t)\n",
        "0"
    },
    table_passed_to_function_by_reference => {
        "local function bump(t) t.n = (t.n or 0) + 1 end\nlocal o = {n = 1}\nbump(o)\nprint(o.n)\n",
        "2"
    },
    iterate_pairs_to_build_key_list_length => {
        "local t = {x=1, y=2, z=3}\nlocal n = 0\nfor _ in pairs(t) do n = n + 1 end\nprint(n)\n",
        "3"
    },
    ipairs_collect_values_into_sum => {
        "local t = {1, 2, 3}\nlocal s = 0\nfor _, v in ipairs(t) do s = s + v end\nprint(s)\n",
        "6"
    },
    default_value_with_or_on_nil_lookup => {
        "local t = {}\nprint(t.missing or 0)\n",
        "0"
    },
    array_index_one_is_first_element_not_zero => {
        "local t = {\"first\", \"second\"}\nprint(t[1])\n",
        "first"
    },
    hash_key_bracket_syntax_for_non_identifier => {
        "local t = {}\nt[\"key-name\"] = 7\nprint(t[\"key-name\"])\n",
        "7"
    },
    length_operator_ignores_hash_keys => {
        "local t = {1, 2, x = 99}\nprint(#t)\n",
        "2"
    },
    table_concat_joins_array_part_only_by_default => {
        "print(table.concat({1, 2, 3}, \",\"))\n",
        "1,2,3"
    },
    insert_at_position_shifts_tail => {
        "local t = {1, 3}\ntable.insert(t, 2, 2)\nprint(table.concat(t, \"\"))\n",
        "123"
    },
    remove_from_middle_compacts_array => {
        "local t = {1, 2, 3}\ntable.remove(t, 2)\nprint(table.concat(t, \"\"))\n",
        "13"
    },
    pack_stores_varargs_in_table_with_n_field => {
        "local p = table.pack(1, 2, 3)\nprint(p.n)\n",
        "3"
    },
    unpack_spreads_list_into_expression_list => {
        "local function sum(a, b, c) return a + b + c end\nprint(sum(table.unpack({1, 2, 3})))\n",
        "6"
    },
    sort_default_orders_numbers_ascending => {
        "local t = {3, 1, 2}\ntable.sort(t)\nprint(t[1] .. t[2] .. t[3])\n",
        "123"
    },
    mixed_table_keeps_both_parts_independent => {
        "local t = {10, key = 20}\nprint(t[1] + t.key)\n",
        "30"
    },
    assign_nil_to_slot_creates_hole_in_array => {
        "local t = {1, 2, 3}\nt[2] = nil\nprint(t[2] == nil)\n",
        "true"
    },
    table_field_dot_syntax_is_sugar_for_brackets => {
        "local t = {name = \"lua\"}\nprint(t[\"name\"])\n",
        "lua"
    },
}
