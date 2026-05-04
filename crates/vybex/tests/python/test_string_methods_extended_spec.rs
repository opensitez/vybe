use super::helpers::*;

macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_python_one($src), $expected);
        }
    };
}

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

compile_case!(str_partition_compile, "parts = 'key=value'.partition('=')\n");
compile_case!(str_rpartition_compile, "parts = 'path/to/file'.rpartition('/')\n");
compile_case!(str_splitlines_compile, "lines = 'a\\nb\\n'.splitlines()\n");
compile_case!(str_splitlines_keepends_compile, "lines = 'a\\nb\\n'.splitlines(True)\n");
runtime_case!(str_isalpha_runtime, "print('Hello'.isalpha())\n", "true");
runtime_case!(str_isalnum_runtime, "print('A1'.isalnum())\n", "true");
runtime_case!(str_isupper_runtime, "print('ABC'.isupper())\n", "true");
runtime_case!(str_islower_runtime, "print('abc'.islower())\n", "true");
runtime_case!(str_isspace_runtime, "print('   '.isspace())\n", "true");
runtime_case!(str_isprintable_runtime, "print('hello'.isprintable())\n", "true");
runtime_case!(str_isidentifier_runtime, "print('snake_case'.isidentifier())\n", "true");
compile_case!(str_istitle_compile, "ok = 'Hello World'.istitle()\n");
compile_case!(str_isdecimal_compile, "ok = '123'.isdecimal()\n");
compile_case!(str_isnumeric_compile, "ok = '123'.isnumeric()\n");
compile_case!(str_casefold_compile, "s = 'Straße'.casefold()\n");
compile_case!(str_expandtabs_compile, "s = 'a\\tb'.expandtabs(4)\n");
compile_case!(str_maketrans_compile, "tbl = str.maketrans({'a': 'x'})\n");
compile_case!(str_translate_compile, "tbl = str.maketrans({'a': 'x'})\ns = 'aba'.translate(tbl)\n");
compile_case!(str_rfind_compile, "i = 'banana'.rfind('na')\n");
compile_case!(str_rindex_compile, "i = 'banana'.rindex('na')\n");
compile_case!(str_format_named_compile, "s = '{greet}, {name}!'.format(greet='Hello', name='Ada')\n");
compile_case!(str_format_positional_compile, "s = '{0} + {0} = {1}'.format(2, 4)\n");
compile_case!(str_percent_format_string_compile, "s = 'hello %s' % 'world'\n");
compile_case!(str_percent_format_tuple_compile, "s = '%s:%d' % ('port', 80)\n");
compile_case!(str_strip_chars_compile, "s = '--hi--'.strip('-')\n");
compile_case!(str_ljust_fill_compile, "s = 'hi'.ljust(5, '.')\n");
compile_case!(str_rjust_fill_compile, "s = 'hi'.rjust(5, '.')\n");
compile_case!(str_center_fill_compile, "s = 'hi'.center(6, '.')\n");
compile_case!(str_encode_utf8_compile, "b = 'hello'.encode('utf-8')\n");
compile_case!(str_split_maxsplit_compile, "parts = 'a,b,c'.split(',', 1)\n");