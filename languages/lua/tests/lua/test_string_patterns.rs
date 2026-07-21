//! Lua patterns — `string.find`, `match`, `gsub`, `gmatch` (Lua 5.x manual §6.4).

lua_print! {
    pattern_digit_class_matches_number => {
        "print(string.match(\"x9y\", \"%d\"))\n",
        "9"
    },
    pattern_letter_class_matches_alpha => {
        "print(string.match(\"1abc2\", \"%a+\"))\n",
        "abc"
    },
    pattern_whitespace_class_matches_space => {
        "print(string.match(\"a b\", \"%s\"))\n",
        " "
    },
    pattern_start_anchor_requires_beginning => {
        "print(string.match(\"lua\", \"^l\"))\n",
        "l"
    },
    pattern_end_anchor_requires_ending => {
        "print(string.match(\"lua\", \"a$\"))\n",
        "a"
    },
    pattern_start_anchor_fails_without_match => {
        "print(tostring(string.match(\"lua\", \"^u\")))\n",
        "nil"
    },
    pattern_capture_returns_substring => {
        "print(string.match(\"id=42\", \"id=(%d+)\"))\n",
        "42"
    },
    pattern_multiple_captures => {
        "print(string.match(\"3+4\", \"(%d+)%+(%d+)\"))\n",
        "3\t4"
    },
    pattern_star_matches_zero_occurrences => {
        "print(string.gsub(\"aaa\", \"a*\", \"b\"))\n",
        "bb\t2"
    },
    pattern_plus_requires_one_or_more => {
        "print(string.gsub(\"aaa\", \"a+\", \"b\"))\n",
        "b\t1"
    },
    pattern_question_makes_optional => {
        "print(string.match(\"colour\", \"colou?r\"))\n",
        "colour"
    },
    pattern_character_class_set => {
        "print(string.match(\"xyz\", \"[xyz]\"))\n",
        "x"
    },
    pattern_negated_character_class => {
        "print(string.match(\"a1\", \"[^%d]\"))\n",
        "a"
    },
    pattern_balanced_parentheses_capture => {
        "print(string.match(\"(())\", \"%b()\"))\n",
        "(())"
    },
    pattern_find_returns_start_and_end => {
        "local s,e=string.find(\"banana\", \"an\")\nprint(s..\",\"..e)\n",
        "2,3"
    },
    pattern_plain_find_disables_magic => {
        "print(string.find(\"a.b\", \".\", 1, true))\n",
        "2\t2"
    },
    pattern_gsub_with_function_replacement => {
        "print(string.gsub(\"a1b2\", \"%d\", function(d) return #d end))\n",
        "a1b1\t2"
    },
    pattern_gmatch_yields_all_words => {
        "local t={}\nfor w in string.gmatch(\"one two\", \"%S+\") do t[#t+1]=w end\nprint(table.concat(t,\",\"))\n",
        "one,two"
    },
    pattern_frontier_followed_by_letter => {
        "print(string.match(\"word\", \"%f[%a]w\"))\n",
        "w"
    },
    pattern_hex_digit_class => {
        "print(string.match(\"gff\", \"%x+\"))\n",
        "ff"
    },
    pattern_punctuation_class => {
        "print(string.match(\"a,b\", \"%p\"))\n",
        ","
    },
    pattern_uppercase_letter_class => {
        "print(string.match(\"aBc\", \"%u\"))\n",
        "B"
    },
    pattern_lowercase_letter_class => {
        "print(string.match(\"aBc\", \"%l\"))\n",
        "a"
    },
    pattern_alphanumeric_class => {
        "print(string.match(\"!ab!\", \"%w+\"))\n",
        "ab"
    },
    pattern_end_of_string_anchor_z => {
        "print(string.match(\"file\\n\", \"%z\"))\n",
        ""
    },
    pattern_non_greedy_minus_suffix => {
        "print(string.match(\"aab\", \"a.-b\"))\n",
        "aab"
    },
    pattern_find_with_start_position_skips_prefix => {
        "print(string.find(\"banana\", \"a\", 3))\n",
        "4\t4"
    },
    pattern_match_returns_nil_on_failure => {
        "print(tostring(string.match(\"abc\", \"z+\")))\n",
        "nil"
    },
    pattern_gsub_limit_replaces_prefix_only => {
        "print(string.gsub(\"aaa\", \"a\", \"b\", 2))\n",
        "bba\t2"
    },
    pattern_caret_inside_class_is_literal => {
        "print(string.match(\"^x\", \"[%^]\"))\n",
        "^"
    },
    pattern_digit_class_matches_numbers => {
        "print(string.match(\"id42\", \"%d+\"))\n",
        "42"
    },
    pattern_word_class_skips_punctuation => {
        "print(string.match(\"!!word!!\", \"%w+\"))\n",
        "word"
    },
    pattern_space_class_matches_whitespace => {
        "print(string.match(\"a b\", \"%s\"))\n",
        " "
    },
    pattern_uppercase_class_matches_capital => {
        "print(string.match(\"xYz\", \"%u\"))\n",
        "Y"
    },
    pattern_frontier_pattern_word_boundary => {
        "print(string.match(\"word\", \"%f[%a]w\"))\n",
        "w"
    },
    pattern_rep_exact_three_literal_a => {
        "print(string.match(\"aaa\", \"aaa\"))\n",
        "aaa"
    },
    pattern_one_or_more_previous_class => {
        "print(string.match(\"aaaa\", \"a+\"))\n",
        "aaaa"
    },
    pattern_zero_or_more_greedy_star => {
        "print(string.match(\"aa\", \"a*\"))\n",
        "aa"
    },
    pattern_character_class_range_az => {
        "print(string.match(\"9z\", \"%a\"))\n",
        "z"
    },
    pattern_optional_suffix_question => {
        "print(string.match(\"color\", \"colou?r\"))\n",
        "color"
    },
    pattern_capture_group_returns_substring => {
        "print(string.match(\"year2024\", \"(%d%d%d%d)\"))\n",
        "2024"
    },
    pattern_multiple_captures_in_order => {
        "local a, b = string.match(\"10-20\", \"(%d+)-(%d+)\")\nprint(a .. \",\" .. b)\n",
        "10,20"
    },
    pattern_gsub_replaces_all_occurrences => {
        "print(string.gsub(\"aaa\", \"a\", \"b\"))\n",
        "bbb\t3"
    },
    pattern_find_on_hello_returns_ll_span => {
        "local s, e = string.find(\"hello\", \"ll\")\nprint(s .. \",\" .. e)\n",
        "3,4"
    },
    pattern_match_start_anchor_caret => {
        "print(string.match(\"lua\", \"^%a+\"))\n",
        "lua"
    },
    pattern_match_end_anchor_dollar => {
        "print(string.match(\"file.lua\", \"%.lua$\"))\n",
        ".lua"
    },
    pattern_percent_escape_literal_dot => {
        "print(string.match(\"a.b\", \"a%.b\"))\n",
        "a.b"
    },
    pattern_gmatch_iterates_all_tokens => {
        "local n = 0\nfor _ in string.gmatch(\"a1b2\", \"%d\") do n = n + 1 end\nprint(n)\n",
        "2"
    },
    pattern_lazy_star_minimal_match => {
        "print(string.match(\"abbc\", \"a.-c\"))\n",
        "abbc"
    },
}
