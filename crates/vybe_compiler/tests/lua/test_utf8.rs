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
}
