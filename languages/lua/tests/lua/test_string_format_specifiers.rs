//! `string.format` specifiers — `%o`, `%x`, `%X`, `%e`, `%g`, `%c`, `%%`, width/precision (Lua 5.x §6.4)

lua_print! {
    format_octal_val => { "print(string.format(\"%o\", 8))\n", "10" },
    format_hex_lower_val => { "print(string.format(\"%x\", 255))\n", "ff" },
    format_hex_upper_val => { "print(string.format(\"%X\", 255))\n", "FF" },
    format_scientific_lower_val => { "print(string.format(\"%e\", 1000))\n", "1.000000e+03" },
    format_scientific_upper_val => { "print(string.format(\"%E\", 1000))\n", "1.000000E+03" },
    format_general_g => { "print(string.format(\"%g\", 100.0))\n", "100" },
    format_general_g_large => { "print(string.format(\"%g\", 1e10))\n", "1e+10" },
    format_char => { "print(string.format(\"%c\", 65))\n", "A" },
    format_percent => { "print(string.format(\"100%%\"))\n", "100%" },
    format_width_right => { "print(string.format(\"%5d\", 42))\n", "   42" },
    format_width_left => { "print(string.format(\"%-5d|\", 42))\n", "42   |" },
    format_zero_pad => { "print(string.format(\"%05d\", 42))\n", "00042" },
    format_plus_sign => { "print(string.format(\"%+d\", 42))\n", "+42" },
    format_precision_f => { "print(string.format(\"%.2f\", 3.14159))\n", "3.14" },
    format_precision_s => { "print(string.format(\"%.3s\", \"hello\"))\n", "hel" },
    format_multi => { "print(string.format(\"%s=%d\", \"x\", 7))\n", "x=7" },
    format_q_escape => { "print(string.format(\"%q\", \"a\\tb\"))\n", "\"a\\tb\"" },
    format_unsigned => { "print(string.format(\"%u\", 42))\n", "42" },
    format_width_str => { "print(string.format(\"%10s|\", \"hi\"))\n", "        hi|" },
}
