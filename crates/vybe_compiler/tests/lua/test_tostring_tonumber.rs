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
}
