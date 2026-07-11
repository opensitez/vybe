//! `string.gsub` with patterns, counts, and anchors (Lua 5.x §6.4.1)

lua_print! {
    gsub_count_vowels => {
        "local _, n = string.gsub(\"hello world\", \"[aeiou]\", \"\")\nprint(n)\n",
        "3"
    },
    gsub_first_only => {
        "print((string.gsub(\"aaa\", \"a\", \"b\", 1)))\n",
        "baa"
    },
    gsub_anchor_start => {
        "print((string.gsub(\"hello\", \"^h\", \"H\")))\n",
        "Hello"
    },
    gsub_anchor_end => {
        "print((string.gsub(\"world!\", \"!$\", \".\")))\n",
        "world."
    },
    gsub_digit_class => {
        "print((string.gsub(\"a1b2c3\", \"%d\", \"N\")))\n",
        "aNbNcN"
    },
    gsub_word_boundary => {
        "print((string.gsub(\"cat and dog\", \"%a+\", \"X\")))\n",
        "X X X"
    },
    gsub_space_chars => {
        "print((string.gsub(\"a b  c\", \"%s+\", \"-\")))\n",
        "a-b-c"
    },
    gsub_fn_skip => {
        "local r = (string.gsub(\"abc\", \"%a\", function(m)\n  if m ~= \"b\" then return m:upper() end\nend))\nprint(r)\n",
        "AbC"
    },
    gsub_replace_ref => {
        "print((string.gsub(\"hello world\", \"(%a+)\", \"[%1]\")))\n",
        "[hello] [world]"
    },
    gsub_multi_caps => {
        "print((string.gsub(\"key=val\", \"(%a+)=(%a+)\", \"%2=%1\")))\n",
        "val=key"
    },
}
