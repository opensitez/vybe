//! utf8 library: advanced character codes, validation, offsets (Lua 5.3+ §6.5)

lua_print! {
    utf8_char_points => {
        "print(utf8.char(65, 0x3B1, 0x1F600))\n",
        "A\u{03B1}\u{1F600}"
    },
    utf8_codepoint_basic => {
        "local s = \"\u{03B1}\"\nprint(utf8.codepoint(s))\n",
        "945"
    },
    utf8_codepoint_range => {
        "local s = \"A\u{03B1}\"\nlocal a, b = utf8.codepoint(s, 1, #s)\nprint(a .. \",\" .. b)\n",
        "65,945"
    },
    utf8_len_unicode => {
        "print(utf8.len(\"A\u{03B1}\u{1F600}\"))\n",
        "3"
    },
    utf8_len_range => {
        "local s = \"A\u{03B1}\u{1F600}\"\nprint(utf8.len(s, 1, 3))\n",
        "2"
    },
    utf8_offset_pos => {
        "local s = \"A\u{03B1}\u{1F600}\"\nprint(utf8.offset(s, 1) .. \",\" .. utf8.offset(s, 2) .. \",\" .. utf8.offset(s, 3))\n",
        "1,2,4"
    },
    utf8_offset_neg => {
        "local s = \"A\u{03B1}\u{1F600}\"\nprint(utf8.offset(s, -1, #s+1))\n",
        "4"
    },
    utf8_codes_iteration_positions_and_codes => {
        "local s = \"\u{03B1}\u{03B2}\"\nlocal r = \"\"\nfor p, c in utf8.codes(s) do r = r .. p .. \":\" .. c .. \" \" end\nprint(r)\n",
        "1:945 3:946 "
    } }
