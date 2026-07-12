//! Advanced print/format string behaviors (Lua 5.x §6.4)

lua_print! {
    format_large_hex => {
        "print(string.format(\"%08X\", 4278190080))\n",
        "FF000000"
    },
    format_float_precision => {
        "print(string.format(\"%.3f\", 1.2345))\n",
        "1.235"
    },
    format_string_padding => {
        "print(string.format(\"%10s\", \"test\"))\n",
        "      test"
    },
}
