//! String pattern matching extended tests — character classes, anchors, frontiers, nested captures (Lua 5.x §6.4.1)

lua_print! {
    pattern_class_alpha => { "print(string.match(\"123a456\", \"%a\"))\n", "a" },
    pattern_class_digit => { "print(string.match(\"abc1def\", \"%d\"))\n", "1" },
    pattern_class_lower => { "print(string.match(\"ABCdEF\", \"%l\"))\n", "d" },
    pattern_class_upper => { "print(string.match(\"abcDef\", \"%u\"))\n", "D" },
    pattern_class_alnum => { "print(string.match(\"!!!a!!!\", \"%w\"))\n", "a" },
    pattern_class_space => { "print(string.match(\"a b\", \"%s\"))\n", " " },
    pattern_class_punctuation => { "print(string.match(\"abc!def\", \"%p\"))\n", "!" },
    pattern_class_cntrl => { "print(string.match(\"a\\nb\", \"%c\"))\n", "\n" },
    pattern_class_hex => { "print(string.match(\"zFz\", \"%x\"))\n", "F" },
    pattern_class_not_digit => { "print(string.match(\"123a456\", \"%D\"))\n", "a" },
    pattern_class_not_space => { "print(string.match(\"  a  \", \"%S\"))\n", "a" },
    pattern_any_char => { "print(string.match(\"abc\", \".\"))\n", "a" },
    pattern_custom_set => { "print(string.match(\"hello\", \"[aeiou]\"))\n", "e" },
    pattern_custom_complement => { "print(string.match(\"hello\", \"[^hello]\"))\n", "nil" },
    pattern_custom_complement_match => { "print(string.match(\"hello!\", \"[^%a]\"))\n", "!" },
    pattern_anchor_start_match => { "print(string.match(\"abc\", \"^a\"))\n", "a" },
    pattern_anchor_start_fail => { "print(tostring(string.match(\"bac\", \"^a\")))\n", "nil" },
    pattern_anchor_end_match => { "print(string.match(\"abc\", \"c$\"))\n", "c" },
    pattern_anchor_end_fail => { "print(tostring(string.match(\"acb\", \"c$\")))\n", "nil" },
    pattern_zero_or_more_greedy => { "print(string.match(\"abbc\", \"ab*c\"))\n", "abbc" },
    pattern_zero_or_more_lazy => { "print(string.match(\"abbc\", \"ab-c\"))\n", "abbc" },
    pattern_one_or_more_greedy => { "print(string.match(\"abbc\", \"ab+c\"))\n", "abbc" },
    pattern_one_or_more_fail => { "print(tostring(string.match(\"ac\", \"ab+c\")))\n", "nil" },
    pattern_zero_or_one_greedy => { "print(string.match(\"abc\", \"ab?c\"))\n", "abc" },
    pattern_zero_or_one_absent => { "print(string.match(\"ac\", \"ab?c\"))\n", "ac" },
    pattern_escape_magic => { "print(string.match(\"a%b\", \"a%%b\"))\n", "a%b" },
    pattern_balanced_parentheses => { "print(string.match(\"a(b(c)d)e\", \"%b()\"))\n", "(b(c)d)" },
    pattern_frontier_word_start => { "print(string.find(\"hello\", \"%f[%a]h\"))\n", "1\t1" },
    pattern_frontier_word_mid => { "print(tostring(string.find(\"ahello\", \"%f[%a]h\")))\n", "nil" },
    pattern_capture_multiple => {
        "local x, y = string.match(\"a=10\", \"(%a+)=(%d+)\")\nprint(x .. \":\" .. y)\n",
        "a:10"
    },
    pattern_capture_index => { "print(string.match(\"abc\", \"b()\"))\n", "3" },
    pattern_nested_captures_values => {
        "local first, second = string.match(\"abc\", \"(a(b)c)\")\nprint(first .. \",\" .. second)\n",
        "abc,b"
    },
}
