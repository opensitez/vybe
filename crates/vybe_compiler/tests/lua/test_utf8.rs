//! UTF-8 library — `utf8.*` (Lua 5.3+ manual §6.4.1).

lua_print! {
    utf8_len_counts_codepoints => {
        "print(utf8.len(\"lua\"))\n",
        "3"
    },
    utf8_char_builds_from_codepoints => {
        "print(utf8.char(108, 117, 97))\n",
        "lua"
    },
    utf8_codes_iterates_codepoints => {
        "local n = 0\nfor _ in utf8.codes(\"ab\") do n = n + 1 end\nprint(n)\n",
        "2"
    },
    utf8_codepoint_reads_first_character => {
        "print(utf8.codepoint(\"λ\"))\n",
        "955"
    },
    utf8_offset_finds_byte_index => {
        "print(utf8.offset(\"aλb\", 2))\n",
        "2"
    },
    utf8_len_returns_nil_on_invalid_sequence => {
        "print(tostring(utf8.len(\"\\255\")))\n",
        "nil"
    },
    utf8_char_rejects_out_of_range_codepoint => {
        "local ok = pcall(function() utf8.char(0x110000) end)\nprint(ok)\n",
        "false"
    },
    utf8_offset_character_at_end => {
        "print(utf8.offset(\"abc\", -1))\n",
        "3"
    },
    utf8_codes_yields_position_and_codepoint => {
        "for p, c in utf8.codes(\"a\") do print(p .. \":\" .. c) end\n",
        "1:97"
    },
    utf8_len_on_empty_string => {
        "print(utf8.len(\"\"))\n",
        "0"
    },
    utf8_char_max_valid_codepoint => {
        "print(utf8.codepoint(utf8.char(0x10FFFF)))\n",
        "1114111"
    },
    utf8_len_with_byte_range => {
        "print(utf8.len(\"aλb\", 2, 4))\n",
        "1"
    },
    utf8_offset_with_negative_n => {
        "print(utf8.offset(\"aλb\", -1))\n",
        "2"
    },
    utf8_offset_with_zero_n_returns_start_of_character => {
        "print(utf8.offset(\"aλb\", 0, 3))\n",
        "2"
    },
    utf8_offset_out_of_bounds_returns_nil => {
        "print(tostring(utf8.offset(\"abc\", 5)))\n",
        "nil"
    },
    utf8_charpattern_matches_single_codepoints => {
        "local count = 0\nfor _ in string.gmatch(\"aλb\", utf8.charpattern) do count = count + 1 end\nprint(count)\n",
        "3"
    },
    utf8_codepoint_multiple_characters => {
        "local c1, c2 = utf8.codepoint(\"ab\", 1, 2)\nprint(c1 .. \",\" .. c2)\n",
        "97,98"
    },
    utf8_codepoint_invalid_range_raises_error => {
        "local ok, err = pcall(function() utf8.codepoint(\"abc\", 5) end)\nprint(ok)\n",
        "false"
    },
}
