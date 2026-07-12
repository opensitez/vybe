//! `string.pack` / `string.unpack` — binary formats (Lua 5.3+ manual §6.4).

lua_print! {
    pack_unpack_signed_byte => {
        "local s=string.pack(\"b\", -1)\nprint(string.unpack(\"b\", s))\n",
        "-1"
    },
    pack_unpack_unsigned_byte => {
        "local s=string.pack(\"B\", 255)\nprint(string.unpack(\"B\", s))\n",
        "255"
    },
    pack_unpack_little_endian_short => {
        "local s=string.pack(\"<h\", 256)\nprint(string.unpack(\"<h\", s))\n",
        "256"
    },
    pack_unpack_big_endian_short => {
        "local s=string.pack(\">h\", 256)\nprint(string.unpack(\">h\", s))\n",
        "256"
    },
    pack_unpack_little_endian_int => {
        "local s=string.pack(\"<i4\", 1000)\nprint(string.unpack(\"<i4\", s))\n",
        "1000"
    },
    pack_unpack_big_endian_int => {
        "local s=string.pack(\">i4\", 1000)\nprint(string.unpack(\">i4\", s))\n",
        "1000"
    },
    pack_unpack_float => {
        "local s=string.pack(\"f\", 2.5)\nprint(string.unpack(\"f\", s) > 2)\n",
        "true"
    },
    pack_unpack_double => {
        "local s=string.pack(\"d\", 1.5)\nprint(string.unpack(\"d\", s))\n",
        "1.5"
    },
    pack_fixed_string_with_c => {
        "local s=string.pack(\"c4\", \"lua!\")\nprint(string.unpack(\"c4\", s))\n",
        "lua!"
    },
    pack_string_with_length_prefix => {
        "local s=string.pack(\"s\", \"hi\")\nprint(string.unpack(\"s\", s))\n",
        "hi"
    },
    packsize_matches_packed_length => {
        "print(string.packsize(\"<i4\") == #string.pack(\"<i4\", 0))\n",
        "true"
    },
    pack_multiple_values_in_order => {
        "local s=string.pack(\"bBi\", 1, 2, 3)\nlocal a,b,c=string.unpack(\"bBi\", s)\nprint(a..\",\"..b..\",\"..c)\n",
        "1,2,3"
    },
    pack_native_endian_equals_sign => {
        "local s=string.pack(\"=i4\", 7)\nprint(string.unpack(\"=i4\", s))\n",
        "7"
    },
    unpack_with_start_index => {
        "local s=string.pack(\"bb\", 1, 2)\nprint(string.unpack(\"b\", s, 2))\n",
        "2"
    },
    pack_zero_terminated_string_z => {
        "local s=string.pack(\"z\", \"end\")\nprint(string.unpack(\"z\", s))\n",
        "end"
    },
}
