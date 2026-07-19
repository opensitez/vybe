//! `tostring` and `tonumber` conversions (Lua 5.x §6.1)

lua_print! {
    tonumber_int_str => { "print(tonumber(\"42\"))\n", "42" },
    tonumber_float_str => { "print(tonumber(\"3.14\"))\n", "3.14" },
    tonumber_hex_str => { "print(tonumber(\"0xff\"))\n", "255" },
    tonumber_base_2 => { "print(tonumber(\"1010\", 2))\n", "10" },
    tonumber_base_16 => { "print(tonumber(\"ff\", 16))\n", "255" },
    tonumber_base_8 => { "print(tonumber(\"77\", 8))\n", "63" },
    tonumber_invalid_nil => { "print(tostring(tonumber(\"abc\")))\n", "nil" },
    tonumber_empty_nil => { "print(tostring(tonumber(\"\")))\n", "nil" },
    tonumber_int_passthrough => { "print(tonumber(42))\n", "42" },
    tonumber_float_passthrough => { "print(tonumber(3.14))\n", "3.14" },
    tonumber_nil_val => { "print(tostring(tonumber(nil)))\n", "nil" },
    tonumber_whitespace => { "print(tonumber(\"  10  \"))\n", "10" },
    tostring_int_val => { "print(tostring(42))\n", "42" },
    tostring_float_val => { "print(tostring(3.0))\n", "3.0" },
    tostring_bool_t => { "print(tostring(true))\n", "true" },
    tostring_bool_f => { "print(tostring(false))\n", "false" },
    tostring_nil_val => { "print(tostring(nil))\n", "nil" },
    tostring_str => { "print(tostring(\"hi\"))\n", "hi" },
    tostring_custom_meta => {
        "local t = setmetatable({}, {__tostring = function() return \"custom\" end})\nprint(tostring(t))\n",
        "custom"
    },
    tonumber_scientific => { "print(tonumber(\"1e3\"))\n", "1000.0" },
    tonumber_negative => { "print(tonumber(\"-99\"))\n", "-99" },
    tonumber_base_36 => {
        "print(tonumber(\"LUA\", 36))\n",
        "28306"
    },
    tonumber_base_out_of_bounds_raises_error => {
        "local ok, err = pcall(function() tonumber(\"10\", 37) end)\nprint(ok)\n",
        "false"
    },
    tonumber_base_under_bounds_raises_error => {
        "local ok, err = pcall(function() tonumber(\"10\", 1) end)\nprint(ok)\n",
        "false"
    },
    tonumber_whitespace_with_base => {
        "print(tonumber(\"  10  \", 16))\n",
        "16"
    },
    tonumber_hex_uppercase => {
        "print(tonumber(\"0XFF\"))\n",
        "255"
    },
    tostring_on_table_returns_string_default => {
        "print(type(tostring({})) == \"string\")\n",
        "true"
    },
    tonumber_on_non_coercible_table_returns_nil => {
        "print(tostring(tonumber({})))\n",
        "nil"
    },
    tonumber_with_invalid_chars_in_base => {
        "print(tostring(tonumber(\"12\", 2)))\n",
        "nil"
    },
}
