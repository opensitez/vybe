//! Extended str methods: search, split, strip, case, format, partition, translate.

crate::runtime_case!(
    str_split_default,
    "print('a b c'.split())\n",
    "['a', 'b', 'c']"
);
crate::runtime_case!(
    str_split_maxsplit,
    "print('a,b,c'.split(',', 1))\n",
    "['a', 'b,c']"
);
crate::runtime_case!(
    str_rsplit,
    "print('a.b.c'.rsplit('.', 1))\n",
    "['a.b', 'c']"
);
crate::runtime_case!(
    str_splitlines,
    "print('a\\nb'.splitlines())\n",
    "['a', 'b']"
);
crate::runtime_case!(
    str_partition,
    "print('key=value'.partition('='))\n",
    "('key', '=', 'value')"
);
crate::runtime_case!(
    str_rpartition,
    "print('a=b=c'.rpartition('='))\n",
    "('a=b', '=', 'c')"
);
crate::runtime_case!(str_strip, "print('  hi  '.strip())\n", "hi");
crate::runtime_case!(str_lstrip, "print('  hi'.lstrip())\n", "hi");
crate::runtime_case!(str_rstrip, "print('hi  '.rstrip())\n", "hi");
crate::runtime_case!(str_strip_chars, "print('xxxhixxx'.strip('x'))\n", "hi");
crate::runtime_case!(
    str_replace_count,
    "print('aaa'.replace('a', 'b', 2))\n",
    "bba"
);
crate::runtime_case!(
    str_replace_all,
    "print('abcabc'.replace('a', 'z'))\n",
    "zbczbc"
);
crate::runtime_case!(str_find_found, "print('hello'.find('ll'))\n", "2");
crate::runtime_case!(str_find_missing, "print('hello'.find('z'))\n", "-1");
crate::runtime_case!(str_rfind, "print('abcbc'.rfind('bc'))\n", "3");
crate::runtime_case!(str_index_found, "print('abc'.index('b'))\n", "1");
crate::runtime_case!(str_count, "print('banana'.count('a'))\n", "3");
crate::runtime_case!(
    str_startswith_true,
    "print('hello'.startswith('he'))\n",
    "True"
);
crate::runtime_case!(
    str_startswith_false,
    "print('hello'.startswith('lo'))\n",
    "False"
);
crate::runtime_case!(str_endswith_true, "print('hello'.endswith('lo'))\n", "True");
crate::runtime_case!(str_upper, "print('AbC'.upper())\n", "ABC");
crate::runtime_case!(str_lower, "print('AbC'.lower())\n", "abc");
crate::runtime_case!(str_title, "print('hello world'.title())\n", "Hello World");
crate::runtime_case!(str_capitalize, "print('hello'.capitalize())\n", "Hello");
crate::runtime_case!(str_swapcase, "print('AbC'.swapcase())\n", "aBc");
crate::runtime_case!(str_casefold, "print('straße'.casefold())\n", "strasse");
crate::runtime_case!(str_center, "print('ab'.center(6, '-'))\n", "--ab--");
crate::runtime_case!(str_ljust, "print('ab'.ljust(5, '.'))\n", "ab...");
crate::runtime_case!(str_rjust, "print('ab'.rjust(5, '.'))\n", "...ab");
crate::runtime_case!(str_zfill, "print('42'.zfill(5))\n", "00042");
crate::runtime_case!(str_join, "print(','.join(['a', 'b', 'c']))\n", "a,b,c");
crate::runtime_case!(
    str_format_positional,
    "print('{}+{}={}'.format(1, 2, 3))\n",
    "1+2=3"
);
crate::runtime_case!(
    str_format_named,
    "print('{a}-{b}'.format(a=1, b=2))\n",
    "1-2"
);
crate::runtime_case!(str_format_index, "print('{0}{1}'.format('x', 'y'))\n", "xy");
crate::runtime_case!(str_isdigit_true, "print('123'.isdigit())\n", "True");
crate::runtime_case!(str_isdigit_false, "print('12a'.isdigit())\n", "False");
crate::runtime_case!(str_isalpha, "print('abc'.isalpha())\n", "True");
crate::runtime_case!(str_isalnum, "print('abc123'.isalnum())\n", "True");
crate::runtime_case!(str_isspace, "print(' \\t'.isspace())\n", "True");
crate::runtime_case!(str_islower, "print('abc'.islower())\n", "True");
crate::runtime_case!(str_isupper, "print('ABC'.isupper())\n", "True");
crate::runtime_case!(str_istitle, "print('Hello'.istitle())\n", "True");
crate::runtime_case!(
    str_removeprefix,
    "print('foobar'.removeprefix('foo'))\n",
    "bar"
);
crate::runtime_case!(
    str_removesuffix,
    "print('foobar'.removesuffix('bar'))\n",
    "foo"
);
crate::runtime_case!(str_expandtabs, "print('a\\tb'.expandtabs(4))\n", "a   b");
crate::runtime_case!(
    str_encode_utf8,
    "print(list('é'.encode('utf-8')))\n",
    "[195, 169]"
);
crate::runtime_case!(
    str_maketrans_translate,
    "print('abc'.translate(str.maketrans('ab', 'xy')))\n",
    "xyc"
);
crate::runtime_case!(
    str_splitlines_keepends,
    "print('a\\n'.splitlines(keepends=True))\n",
    "['a\\n']"
);
crate::runtime_case!(str_rindex, "print('abcbc'.rindex('bc'))\n", "3");
crate::runtime_case!(str_format_padding, "print('{:05d}'.format(42))\n", "00042");

crate::compile_case!(str_format_spec_float, "print('{:.2f}'.format(3.14159))\n");
crate::compile_case!(str_format_spec_hex, "print('{:#x}'.format(255))\n");
crate::compile_case!(str_format_spec_align, "print('{:<5}'.format('x'))\n");
crate::compile_case!(str_encode_ascii, "s = 'hello'.encode('ascii')\n");
crate::compile_case!(str_decode_bytes, "b'hi'.decode('utf-8')\n");
