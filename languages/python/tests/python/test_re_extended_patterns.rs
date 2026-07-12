//! re module extended: compile flags, groups, split, sub, finditer, escape.

crate::runtime_case!(
    re_search_group,
    "import re\nm = re.search(r'(\\d+)', 'a123b')\nprint(m.group(1))\n",
    "123"
);
crate::runtime_case!(
    re_match_start,
    "import re\nprint(re.match(r'ab', 'abc') is not None)\n",
    "True"
);
crate::runtime_case!(
    re_fullmatch,
    "import re\nprint(re.fullmatch(r'\\d+', '123') is not None)\n",
    "True"
);
crate::runtime_case!(
    re_split,
    "import re\nprint(re.split(r'[,;]', 'a,b;c'))\n",
    "['a', 'b', 'c']"
);
crate::runtime_case!(
    re_sub,
    "import re\nprint(re.sub(r'\\d', 'X', 'a1b2'))\n",
    "aXbX"
);
crate::runtime_case!(
    re_subn,
    "import re\nprint(re.subn(r'a', 'b', 'aba')[1])\n",
    "2"
);
crate::runtime_case!(
    re_findall,
    "import re\nprint(re.findall(r'\\w+', 'hi there'))\n",
    "['hi', 'there']"
);
crate::runtime_case!(
    re_finditer_count,
    "import re\nprint(len(list(re.finditer(r'\\d', 'a1b2c3'))))\n",
    "3"
);
crate::runtime_case!(
    re_compile_pattern,
    "import re\np = re.compile(r'[a-z]+')\nprint(p.findall('Ab cd'))\n",
    "['b', 'cd']"
);
crate::runtime_case!(
    re_escape_literal,
    "import re\nprint(re.escape('a.b') in re.escape('a.b'))\n",
    "True"
);
crate::runtime_case!(
    re_groups_multiple,
    "import re\nm = re.match(r'(a)(b)', 'ab')\nprint(m.groups())\n",
    "('a', 'b')"
);
crate::runtime_case!(
    re_groupdict_named,
    "import re\nm = re.match(r'(?P<x>a)(?P<y>b)', 'ab')\nprint(m.groupdict()['y'])\n",
    "b"
);
crate::runtime_case!(
    re_span,
    "import re\nm = re.search('bc', 'abc')\nprint(m.span())\n",
    "(1, 3)"
);
crate::runtime_case!(
    re_start_end,
    "import re\nm = re.search('bc', 'abc')\nprint(m.start(), m.end())\n",
    "1 3"
);
crate::runtime_case!(
    re_ignorecase,
    "import re\nprint(re.findall('a+', 'AaA', re.I))\n",
    "['AaA']"
);
crate::runtime_case!(
    re_multiline_caret,
    "import re\nprint(re.findall('^a', 'a\\na', re.M))\n",
    "['a', 'a']"
);
crate::runtime_case!(
    re_dotall,
    "import re\nprint(re.findall('a.b', 'a\\nb', re.S))\n",
    "['a\\nb']"
);
crate::runtime_case!(
    re_non_greedy,
    "import re\nprint(re.findall('<.*?>', '<a><b>'))\n",
    "['<a>', '<b>']"
);
crate::runtime_case!(
    re_digit_class,
    "import re\nprint(re.findall(r'\\d+', 'a12b34'))\n",
    "['12', '34']"
);
crate::runtime_case!(
    re_word_boundary,
    "import re\nprint(re.findall(r'\\bword\\b', 'a word here'))\n",
    "['word']"
);
crate::runtime_case!(
    re_alternation,
    "import re\nprint(re.findall(r'cat|dog', 'cat dog'))\n",
    "['cat', 'dog']"
);
crate::runtime_case!(
    re_quantifier_exact,
    "import re\nprint(re.findall(r'a{3}', 'aa aaa aaaa'))\n",
    "['aaa', 'aaa']"
);
crate::runtime_case!(
    re_optional,
    "import re\nprint(re.findall(r'colou?r', 'color colour'))\n",
    "['color', 'colour']"
);
crate::runtime_case!(
    re_lookahead,
    "import re\nprint(re.findall(r'(?=\\d)\\d+', 'a1b23'))\n",
    "['1', '23']"
);
crate::runtime_case!(
    re_negative_lookahead,
    "import re\nprint(re.findall(r'\\d+(?!\\d)', '123 45'))\n",
    "['3', '5']"
);
crate::runtime_case!(
    re_comment_flag,
    "import re\nprint(re.findall(r'a+ # vowel', 'aaa', re.X))\n",
    "['aaa']"
);
crate::runtime_case!(
    re_split_maxsplit,
    "import re\nprint(re.split(r'\\s+', 'a  b   c', maxsplit=1))\n",
    "['a', 'b   c']"
);
crate::runtime_case!(
    re_sub_count,
    "import re\nprint(re.sub(r'x', 'y', 'xxx'))\n",
    "yyy"
);
crate::runtime_case!(
    re_pattern_flags,
    "import re\np = re.compile('abc', re.IGNORECASE)\nprint(p.findall('aBc AbC'))\n",
    "['aBc', 'AbC']"
);
crate::runtime_case!(
    re_match_none,
    "import re\nprint(re.match('z', 'abc') is None)\n",
    "True"
);
crate::runtime_case!(
    re_search_none,
    "import re\nprint(re.search('z', 'abc') is None)\n",
    "True"
);
crate::runtime_case!(
    re_findall_empty,
    "import re\nprint(re.findall('z', 'abc'))\n",
    "[]"
);
crate::runtime_case!(
    re_group_zero,
    "import re\nm = re.search('(hi)', 'hi')\nprint(m.group(0))\n",
    "hi"
);
crate::runtime_case!(
    re_lastindex,
    "import re\nm = re.match('(a)(b)', 'ab')\nprint(m.lastindex)\n",
    "2"
);
crate::runtime_case!(
    re_pattern_sub,
    "import re\np = re.compile(r'\\s+')\nprint(p.sub('-', 'a b c'))\n",
    "a-b-c"
);
crate::runtime_case!(
    re_pattern_split,
    "import re\np = re.compile(r',')\nprint(p.split('a,b'))\n",
    "['a', 'b']"
);
crate::runtime_case!(
    re_ascii_flag,
    "import re\nprint(re.findall(r'\\w+', 'café', re.A))\n",
    "[]"
);
crate::runtime_case!(
    re_verbose,
    "import re\np = re.compile(r'''\\d+  # digits''', re.X)\nprint(p.findall('a12'))\n",
    "['12']"
);
crate::runtime_case!(
    re_backreference,
    "import re\nprint(re.findall(r'(.)\\1', 'aabcc'))\n",
    "['aa', 'cc']"
);
crate::runtime_case!(
    re_non_capturing,
    "import re\nm = re.match(r'(?:ab)(c)', 'abc')\nprint(m.group(1))\n",
    "c"
);
crate::runtime_case!(
    re_unicode_word,
    "import re\nprint(len(re.findall(r'\\w+', 'hello')))\n",
    "1"
);
crate::runtime_case!(
    re_search_pos,
    "import re\nprint(re.search('b', 'abc', pos=1).group())\n",
    "b"
);
crate::runtime_case!(
    re_match_endpos,
    "import re\nprint(re.match('abc', 'abcxyz', endpos=3) is not None)\n",
    "True"
);
crate::runtime_case!(
    re_finditer_empty_match,
    "import re\nprint(len(list(re.finditer(r'a?', 'bbb'))))\n",
    "4"
);
crate::runtime_case!(
    re_sub_function,
    "import re\nprint(re.sub(r'\\d', lambda m: 'X', 'a1'))\n",
    "aX"
);
crate::runtime_case!(
    re_compile_cache,
    "import re\np1 = re.compile('a')\np2 = re.compile('a')\nprint(p1.pattern)\n",
    "a"
);
crate::runtime_case!(
    re_error_invalid,
    "import re\ntry:\n re.compile('(')\n print('ok')\nexcept re.error:\n print('bad')\n",
    "bad"
);

crate::compile_case!(
    re_template_flag,
    "import re\nre.findall('(a)', 'aba', re.T)\n"
);
crate::compile_case!(re_debug_flag, "import re\nre.compile('a', re.DEBUG)\n");
crate::compile_case!(re_locale_flag, "import re\nre.compile(r'\\w', re.L)\n");
crate::compile_case!(re_unicode_flag, "import re\nre.compile(r'\\w', re.U)\n");
crate::compile_case!(re_scanner, "import re\nre.Scanner([('a', 'A')])\n");
