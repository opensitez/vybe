//! Strings — `#`, `..`, and string library (Lua 5.x manual §3.4.6–3.4.7, §6.4).

lua_print! {
    length_of_ascii_string => { "print(#\"hello\")\n", "5" },
    length_of_empty_string => { "print(#\"\")\n", "0" },
    concat_coerces_numbers => { "print(1 .. 2 .. 3)\n", "123" },
    string_len_builtin => { "print(string.len(\"lua\"))\n", "3" },
    string_sub_extracts_slice => { "print(string.sub(\"hello\", 2, 4))\n", "ell" },
    string_upper_case => { "print(string.upper(\"AbC\"))\n", "ABC" },
    string_lower_case => { "print(string.lower(\"AbC\"))\n", "abc" },
    string_find_returns_start_index => { "print(string.find(\"banana\", \"ana\"))\n", "2\t4" },
    string_find_returns_nil_when_absent => {
        "print(tostring(string.find(\"abc\", \"z\")))\n",
        "nil"
    },
    string_find_plain_search_fourth_arg => {
        "print(string.find(\"a%%b\", \"%\", 1, true))\n",
        "2\t2"
    },
    string_sub_with_negative_end_index => {
        "print(string.sub(\"hello\", 1, -2))\n",
        "hell"
    },
    string_sub_from_position_to_end => {
        "print(string.sub(\"hello\", 3))\n",
        "llo"
    },
    string_rep_repeats_pattern => { "print(string.rep(\"ab\", 3))\n", "ababab" },
    string_rep_with_separator => { "print(string.rep(\"x\", 3, \",\"))\n", "x,x,x" },
    string_reverse_flips_order => { "print(string.reverse(\"lua\"))\n", "aul" },
    string_byte_of_first_char => { "print(string.byte(\"A\"))\n", "65" },
    string_byte_with_position => { "print(string.byte(\"ABC\", 2))\n", "66" },
    string_char_builds_from_codes => { "print(string.char(72, 105))\n", "Hi" },
    string_format_integer_placeholder => { "print(string.format(\"%d\", 42))\n", "42" },
    string_format_string_and_number => { "print(string.format(\"%s=%d\", \"port\", 80))\n", "port=80" },
    string_format_quoting_escape => { "print(string.format(\"%q\", \"a\\n\"))\n", "\"a\\n\"" },
    string_match_anchor_start => { "print(string.match(\"hello\", \"^h\"))\n", "h" },
    string_match_capture_group => {
        "print(string.match(\"user=ada\", \"user=(.+)\") or \"\")\n",
        "ada"
    },
    string_gsub_single_replacement => {
        "print(string.gsub(\"banana\", \"a\", \"A\", 1))\n",
        "bAnana\t1"
    },
    string_gsub_returns_replacement_count => {
        "local _,n=string.gsub(\"aaa\", \"a\", \"b\")\nprint(n)\n",
        "3"
    },
    string_gmatch_iterates_tokens => {
        "local n=0\nfor _ in string.gmatch(\"one two three\", \"%S+\") do n=n+1 end\nprint(n)\n",
        "3"
    },
    lexicographic_order_is_byte_based => { "print(\"A\" < \"a\")\n", "true" },
    empty_string_concatenation_is_identity => { "print(\"\" .. \"z\")\n", "z" },
    string_find_returns_capture_start_and_end => {
        "local s,e=string.find(\"banana\", \"an\")\nprint(s .. \",\" .. e)\n",
        "2,3"
    },
    string_match_returns_nil_without_match => {
        "print(tostring(string.match(\"abc\", \"z\")))\n",
        "nil"
    },
    string_gsub_with_table_replacement => {
        "print(string.gsub(\"a b\", \"%a\", {a=\"A\", b=\"B\"}))\n",
        "A B\t2"
    },
    string_format_hex_uppercase => { "print(string.format(\"%X\", 255))\n", "FF" },
    string_format_float_precision => { "print(string.format(\"%.1f\", 3.14))\n", "3.1" },
    string_byte_range_returns_varargs => {
        "local a,b=string.byte(\"ABC\", 1, 2)\nprint(a .. \",\" .. b)\n",
        "65,66"
    },
    string_pack_unpack_roundtrip_integer => {
        "local s=string.pack(\">i4\", 42)\nprint(string.unpack(\">i4\", s))\n",
        "42\t5"
    },
    string_packsize_reports_byte_count => {
        "print(string.packsize(\">i4\"))\n",
        "4"
    },
    string_find_honors_start_position => {
        "print(string.find(\"banana\", \"a\", 3))\n",
        "4\t4"
    },
    string_match_end_anchor => {
        "print(string.match(\"end\", \"d$\"))\n",
        "d"
    },
    string_rep_zero_times_is_empty => {
        "print(string.rep(\"x\", 0))\n",
        ""
    },
    string_reverse_empty_string => {
        "print(string.reverse(\"\"))\n",
        ""
    },
    string_format_character_placeholder => {
        "print(string.format(\"%c\", 65))\n",
        "A"
    },
    string_gsub_replaces_all_when_limit_omitted => {
        "print(string.gsub(\"aaa\", \"a\", \"b\"))\n",
        "bbb\t3"
    },
    string_gmatch_capture_groups_iterate => {
        "local s = \"\"\nfor a,b in string.gmatch(\"1=2\", \"(%d+)=(%d+)\") do s = a .. \"+\" .. b end\nprint(s)\n",
        "1+2"
    },
    hash_length_matches_string_len_builtin => {
        "print(#\"hello\" == string.len(\"hello\"))\n",
        "true"
    },
    gsub_removes_spaces_common_cleanup => {
        "print(string.gsub(\" a b \", \"%s+\", \" \"))\n",
        " a b \t3"
    },
    sub_checks_prefix_before_processing => {
        "local s = \"lua-5.4\"\nprint(string.sub(s, 1, 3) == \"lua\")\n",
        "true"
    },
    upper_and_lower_for_case_insensitive_compare => {
        "print(string.lower(\"ABC\") == string.lower(\"abc\"))\n",
        "true"
    },
    format_builds_simple_message => {
        "print(string.format(\"hello %s\", \"world\"))\n",
        "hello world"
    },
    join_path_segments_with_concat => {
        "print(\"dir\" .. \"/\" .. \"file.lua\")\n",
        "dir/file.lua"
    },
    strip_leading_character_with_sub => {
        "print(string.sub(\"#tag\", 2))\n",
        "tag"
    },
    test_empty_string_is_truthy_in_or => {
        "print((\"\" or \"fallback\") == \"\")\n",
        "true"
    },
    build_csv_line_from_array => {
        "local t = {\"a\", \"b\", \"c\"}\nprint(table.concat(t, \",\"))\n",
        "a,b,c"
    },
    count_char_occurrence_manual_loop => {
        "local s = \"banana\"\nlocal n = 0\nfor i = 1, #s do if string.sub(s,i,i) == \"a\" then n = n + 1 end end\nprint(n)\n",
        "3"
    },
    replace_first_space_with_dash => {
        "print(string.gsub(\"a b c\", \" \", \"-\", 1))\n",
        "a-b c\t1"
    },
    extract_file_extension_with_match => {
        "print(string.match(\"file.lua\", \"%.(%w+)$\"))\n",
        "lua"
    },
    pad_left_with_rep_and_concat => {
        "local s = \"7\"\nprint(string.rep(\"0\", 3 - #s) .. s)\n",
        "007"
    },
    compare_strings_case_insensitive_via_lower => {
        "print(string.lower(\"ABC\") == string.lower(\"abc\"))\n",
        "true"
    },
    split_lines_on_newline_with_gmatch => {
        "local n = 0\nfor _ in string.gmatch(\"a\\nb\\n\", \"[^\\n]+\") do n = n + 1 end\nprint(n)\n",
        "2"
    },
    string_is_empty_when_length_zero => {
        "print(#\"\" == 0)\n",
        "true"
    },
    concat_number_and_string_for_log_line => {
        "print(\"value=\" .. 42)\n",
        "value=42"
    },
    find_substring_position_or_negative_sentinel => {
        "local i = string.find(\"hello\", \"z\")\nprint(i == nil)\n",
        "true"
    },
    reverse_string_for_palindrome_check_setup => {
        "print(string.reverse(\"stressed\"))\n",
        "desserts"
    },
    concat_does_not_mutate_original_string => {
        "local s = \"ab\"\nlocal t = s .. \"c\"\nprint(s .. \",\" .. t)\n",
        "ab,abc"
    },
    string_length_unchanged_after_concat_into_new => {
        "local base = \"x\"\nlocal extended = base .. \"y\"\nprint(#base)\n",
        "1"
    },
    repeated_concat_builds_greeting => {
        "local name = \"lua\"\nprint(\"hello \" .. name .. \"!\")\n",
        "hello lua!"
    },
    tonumber_on_numeric_string_for_concat => {
        "local n = tonumber(\"42\")\nprint(\"n=\" .. n)\n",
        "n=42"
    },
    string_byte_returns_first_char_code => {
        "print(string.byte(\"A\"))\n",
        "65"
    },
    string_char_from_code_point => {
        "print(string.char(66))\n",
        "B"
    },
    string_rep_duplicates_pattern => {
        "print(string.rep(\"ab\", 3))\n",
        "ababab"
    },
    sub_with_single_index_returns_one_char => {
        "print(string.sub(\"hello\", 1, 1))\n",
        "h"
    },
    empty_string_concat_is_identity_on_other_side => {
        "print(\"x\" .. \"\")\n",
        "x"
    },
    string_equality_is_by_value => {
        "print(\"a\" == \"a\")\n",
        "true"
    },
    string_inequality_detects_different_content => {
        "print(\"a\" ~= \"b\")\n",
        "true"
    },
    string_format_left_align => {
        "print(string.format(\"%-6s\", \"abc\") .. \"*\")\n",
        "abc   *"
    },
    string_format_scientific_notation => {
        "print(string.format(\"%.1e\", 123.45))\n",
        "1.2e+02"
    },
    string_format_sign_forces_plus => {
        "print(string.format(\"%+d\", 42) .. \" \" .. string.format(\"%+d\", -42))\n",
        "+42 -42"
    },
    string_format_space_prefix => {
        "print(string.format(\"% d\", 42) .. \"|\" .. string.format(\"% d\", -42))\n",
        " 42|-42"
    },
    string_format_hex_alternative_form => {
        "print(string.format(\"%#x\", 255))\n",
        "0xff"
    },
    string_format_zero_padded_integer => {
        "print(string.format(\"%04d\", 7))\n",
        "0007"
    },
    string_format_string_precision_truncation => {
        "print(string.format(\"%.3s\", \"hello\"))\n",
        "hel"
    },
    string_pack_float_roundtrip => {
        "local s = string.pack(\"f\", 1.25)\nprint(string.unpack(\"f\", s))\n",
        "1.25\t5"
    },
    string_pack_double_roundtrip => {
        "local s = string.pack(\"d\", 1.125)\nprint(string.unpack(\"d\", s))\n",
        "1.125\t9"
    },
    string_pack_little_endian_i2 => {
        "local s = string.pack(\"<i2\", -1000)\nprint(string.unpack(\"<i2\", s))\n",
        "-1000\t3"
    },
    string_pack_big_endian_I2 => {
        "local s = string.pack(\">I2\", 1000)\nprint(string.unpack(\">I2\", s))\n",
        "1000\t3"
    },
    string_pack_zero_terminated_string => {
        "local s = string.pack(\"z\", \"abc\")\nprint(#s .. \" \" .. string.unpack(\"z\", s))\n",
        "4 abc"
    },
    string_packsize_complex_format => {
        "print(string.packsize(\"i4 b h\"))\n",
        "7"
    },
    string_gsub_function_retains_nil_results => {
        "local res = string.gsub(\"a b c\", \"%a\", function(x) if x == \"b\" then return \"B\" end end)\nprint(res)\n",
        "a B c"
    },
    string_gsub_table_ignores_missing_keys => {
        "local res = string.gsub(\"a b c\", \"%a\", {a=\"A\"})\nprint(res)\n",
        "A b c"
    },
    string_gsub_with_capture_references => {
        "print(string.gsub(\"10-20\", \"(%d+)-(%d+)\", \"%2/%1\"))\n",
        "20/10\t1"
    } }
