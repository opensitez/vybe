//! String library exhaustive tests: sub, len, reverse, lower, upper, rep, formats (Lua 5.x §6.4)

lua_print! {
    str_exh_sub_basic => { "print(string.sub(\"hello\", 2, 4))\n", "ell" },
    str_exh_sub_negative => { "print(string.sub(\"hello\", -3, -1))\n", "llo" },
    str_exh_sub_out_bounds => { "print(string.sub(\"hello\", 1, 10))\n", "hello" },
    str_exh_len => { "print(string.len(\"hello\"))\n", "5" },
    str_exh_reverse => { "print(string.reverse(\"hello\"))\n", "olleh" },
    str_exh_lower => { "print(string.lower(\"HELLO\"))\n", "hello" },
    str_exh_upper => { "print(string.upper(\"hello\"))\n", "HELLO" },
    str_exh_rep => { "print(string.rep(\"ab\", 3))\n", "ababab" },
    str_exh_rep_sep => { "print(string.rep(\"ab\", 3, \"-\"))\n", "ab-ab-ab" },
    str_exh_format_percent => { "print(string.format(\"%%\", 10))\n", "%" },
    str_exh_format_string => { "print(string.format(\"%s\", \"hello\"))\n", "hello" },
    str_exh_format_decimal => { "print(string.format(\"%d\", 42))\n", "42" },
    str_exh_format_unsigned => { "print(string.format(\"%u\", 42))\n", "42" },
    str_exh_format_octal => { "print(string.format(\"%o\", 8))\n", "10" },
    str_exh_format_hex_lower => { "print(string.format(\"%x\", 255))\n", "ff" },
    str_exh_format_hex_upper => { "print(string.format(\"%X\", 255))\n", "FF" },
    str_exh_format_float => { "print(string.format(\"%.2f\", 3.1415))\n", "3.14" },
}
