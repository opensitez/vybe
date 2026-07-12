//! Coercion — tonumber / tostring (Lua 5.x manual §3.4.2).

lua_print! {
    tonumber_parses_integer_string => { "print(tonumber(\"42\"))\n", "42" },
    tonumber_parses_float_string => { "print(tonumber(\"3.5\"))\n", "3.5" },
    tonumber_invalid_yields_nil => { "print(tostring(tonumber(\"xyz\")))\n", "nil" },
    tostring_number => { "print(tostring(123))\n", "123" },
    tostring_boolean_true => { "print(tostring(true))\n", "true" },
    tostring_nil => { "print(tostring(nil))\n", "nil" },
    tostring_false => { "print(tostring(false))\n", "false" },
    tonumber_parses_hex_with_base_sixteen => { "print(tonumber(\"ff\", 16))\n", "255" },
    tonumber_parses_binary_with_base_two => { "print(tonumber(\"1010\", 2))\n", "10" },
    tonumber_parses_octal_with_base_eight => { "print(tonumber(\"17\", 8))\n", "15" },
    tonumber_stops_at_invalid_suffix => { "print(tostring(tonumber(\"42px\")))\n", "nil" },
    tonumber_leading_spaces_allowed => { "print(tonumber(\"  99\"))\n", "99" },
    arithmetic_coerces_string_to_number => { "print(\"5\" + 3)\n", "8" },
    concatenation_coerces_number_to_string => { "print(\"v\" .. 3)\n", "v3" },
    tonumber_for_input_validation => {
        "local input = \"42\"\nlocal n = tonumber(input)\nprint(n ~= nil)\n",
        "true"
    },
    tonumber_rejects_non_numeric_user_input => {
        "local input = \"abc\"\nprint(tostring(tonumber(input)))\n",
        "nil"
    },
    tostring_for_building_messages => {
        "local n = 7\nprint(\"count=\" .. tostring(n))\n",
        "count=7"
    },
    compare_after_tonumber_conversion => {
        "print(tonumber(\"10\") > 5)\n",
        "true"
    },
    concatenate_after_tostring_on_boolean => {
        "print(\"ok=\" .. tostring(true))\n",
        "ok=true"
    },
    tonumber_on_negative_decimal_string => {
        "print(tonumber(\"-2.5\"))\n",
        "-2.5"
    },
    tonumber_returns_nil_for_empty_string => {
        "print(tostring(tonumber(\"\")))\n",
        "nil"
    },
    add_after_tonumber_from_input => {
        "print(tonumber(\"3\") + 4)\n",
        "7"
    },
    tostring_on_table_without_metamethod_is_not_nil => {
        "print(type(tostring({})))\n",
        "string"
    },
}
