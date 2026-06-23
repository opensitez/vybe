//! `string.format` — placeholders, flags, width, precision (Lua 5.x manual §6.4).

lua_print! {
    format_percent_literal_doubles_percent => {
        "print(string.format(\"%%\"))\n",
        "%"
    },
    format_decimal_integer => { "print(string.format(\"%d\", 42))\n", "42" },
    format_signed_integer_with_plus_flag => { "print(string.format(\"%+d\", 7))\n", "+7" },
    format_signed_integer_with_space_flag => { "print(string.format(\"% d\", 7))\n", " 7" },
    format_unsigned_integer => { "print(string.format(\"%u\", 255))\n", "255" },
    format_octal_integer => { "print(string.format(\"%o\", 8))\n", "10" },
    format_hex_lowercase => { "print(string.format(\"%x\", 255))\n", "ff" },
    format_hex_uppercase => { "print(string.format(\"%X\", 255))\n", "FF" },
    format_float_default_six_decimals => { "print(string.format(\"%f\", 1))\n", "1.000000" },
    format_float_with_precision => { "print(string.format(\"%.2f\", 3.1415))\n", "3.14" },
    format_scientific_lowercase => { "print(string.format(\"%e\", 1000))\n", "1.000000e+03" },
    format_scientific_uppercase => { "print(string.format(\"%E\", 1000))\n", "1.000000E+03" },
    format_general_shortest_repr => { "print(string.format(\"%g\", 3.14))\n", "3.14" },
    format_character_from_code => { "print(string.format(\"%c\", 66))\n", "B" },
    format_string_placeholder => { "print(string.format(\"%s\", \"lua\"))\n", "lua" },
    format_quoted_string_escapes => { "print(string.format(\"%q\", \"a\"))\n", "\"a\"" },
    format_width_pads_string_right => { "print(string.format(\"%5s\", \"x\"))\n", "    x" },
    format_minus_flag_left_justifies => { "print(string.format(\"%-5s\", \"x\"))\n", "x    " },
    format_zero_flag_pads_number => { "print(string.format(\"%05d\", 7))\n", "00007" },
    format_hash_flag_adds_decimal_for_float => {
        "print(string.format(\"%#g\", 3.0))\n",
        "3.0"
    },
    format_positional_arguments => {
        "print(string.format(\"%2$s %1$d\", 9, \"ok\"))\n",
        "ok 9"
    },
    format_multiple_values_in_one_call => {
        "print(string.format(\"%d %s\", 1, \"two\"))\n",
        "1 two"
    },
    format_negative_integer => { "print(string.format(\"%d\", -5))\n", "-5" },
    format_zero_integer => { "print(string.format(\"%d\", 0))\n", "0" },
    format_large_width_on_integer => { "print(string.format(\"%6d\", 12))\n", "    12" },
    format_precision_truncates_string => { "print(string.format(\"%.2s\", \"hello\"))\n", "he" },
    format_hex_with_width_and_zero_pad => { "print(string.format(\"%04x\", 15))\n", "000f" },
    format_string_with_embedded_percent => {
        "print(string.format(\"100%% done\"))\n",
        "100% done"
    },
    format_float_negative_zero_shows_sign_with_plus => {
        "print(string.format(\"%+f\", -0.0) == \"-0.000000\" or string.format(\"%+f\", -0.0) == \"+0.000000\")\n",
        "true"
    },
    format_concatenates_literal_and_placeholder => {
        "print(string.format(\"n=%d\", 3))\n",
        "n=3"
    },
}
