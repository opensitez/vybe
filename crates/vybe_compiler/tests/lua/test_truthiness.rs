//! Types and truthiness — `type()`, `nil` vs `false`, Lua truth rules (§3.3.1).

lua_print! {
    nil_in_if_else_takes_else_branch => {
        "if nil then print(\"a\") else print(\"b\") end\n",
        "b"
    },
    false_in_if_else_takes_else_branch => {
        "if false then print(\"a\") else print(\"b\") end\n",
        "b"
    },
    empty_table_is_truthy_in_if => {
        "if {} then print(\"yes\") else print(\"no\") end\n",
        "yes"
    },
    empty_string_is_truthy_in_or_expression => {
        "print((\"\" or \"fallback\") == \"\")\n",
        "true"
    },
    zero_is_truthy_in_or_expression => {
        "print((0 or 99) == 0)\n",
        "true"
    },
    type_distinguishes_nil_from_boolean => {
        "print(type(nil) ~= type(false))\n",
        "true"
    },
    type_distinguishes_number_from_string => {
        "print(type(1) ~= type(\"1\"))\n",
        "true"
    },
    and_short_circuit_skips_right_when_left_false => {
        "print(false and print(\"rhs\"))\n",
        "false"
    },
    or_short_circuit_skips_right_when_left_true => {
        "print(true or print(\"rhs\"))\n",
        "true"
    },
    not_nil_is_true => {
        "print(not nil)\n",
        "true"
    },
    not_zero_is_false => {
        "print(not 0)\n",
        "false"
    },
    not_empty_string_is_false => {
        "print(not \"\")\n",
        "false"
    },
    nil_equals_only_nil => {
        "print(nil == nil)\n",
        "true"
    },
    false_equals_only_false => {
        "print(false == false)\n",
        "true"
    },
    if_checks_variable_against_nil_explicitly => {
        "local v = nil\nif v == nil then print(\"unset\") end\n",
        "unset"
    },
    type_nil_is_distinct_from_type_boolean => {
        "print(type(nil) .. \",\" .. type(false))\n",
        "nil,boolean"
    },
    type_number_for_integer_literal => {
        "print(type(0))\n",
        "number"
    },
    type_string_for_quoted_literal => {
        "print(type(\"\"))\n",
        "string"
    },
    type_table_for_brace_constructor => {
        "print(type({}))\n",
        "table"
    },
    only_false_and_nil_are_falsy_in_not => {
        "print((not false) and (not nil) and (not not 0) and (not not \"\"))\n",
        "true"
    },
    boolean_true_is_truthy_in_if => {
        "if true then print(\"yes\") end\n",
        "yes"
    },
    table_with_false_field_value_is_still_truthy => {
        "local t = {flag = false}\nif t then print(\"table\") end\n",
        "table"
    },
}
