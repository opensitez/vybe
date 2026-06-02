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

runtime_case!(
    dict_get_existing_runtime,
    "d = {'a': 1, 'b': 2}\nprint(d.get('a'))\n",
    "1"
);
runtime_case!(
    dict_get_missing_with_default_runtime,
    "d = {'a': 1}\nprint(d.get('z', 99))\n",
    "99"
);
runtime_case!(
    dict_get_missing_without_default_runtime,
    "d = {'a': 1}\nprint(d.get('z'))\n",
    "null"
);
runtime_case!(
    dict_setdefault_inserts_runtime,
    "d = {}\nprint(d.setdefault('x', 7))\n",
    "7"
);
runtime_case!(
    dict_setdefault_preserves_runtime,
    "d = {'x': 3}\nprint(d.setdefault('x', 7))\n",
    "3"
);
runtime_case!(
    dict_update_literal_runtime,
    "d = {'a': 1}\nd.update({'b': 2})\nprint(d['b'])\n",
    "2"
);
runtime_case!(
    dict_update_overwrites_runtime,
    "d = {'a': 1}\nd.update({'a': 9})\nprint(d['a'])\n",
    "9"
);
runtime_case!(
    dict_copy_independent_runtime,
    "a = {'x': 1}\nb = a.copy()\nb['x'] = 8\nprint(a['x'])\n",
    "1"
);
runtime_case!(
    dict_clear_runtime,
    "d = {'a': 1, 'b': 2}\nd.clear()\nprint(len(d))\n",
    "0"
);
runtime_case!(
    dict_pop_existing_runtime,
    "d = {'a': 1, 'b': 2}\nprint(d.pop('a'))\n",
    "1"
);
runtime_case!(
    dict_pop_missing_default_runtime,
    "d = {'a': 1}\nprint(d.pop('z', 5))\n",
    "5"
);
runtime_case!(
    dict_constructor_keyword_runtime,
    "d = dict(alpha=1, beta=2)\nprint(d['beta'])\n",
    "2"
);
runtime_case!(
    dict_constructor_pairs_runtime,
    "d = dict([('a', 1), ('b', 2)])\nprint(d['b'])\n",
    "2"
);
runtime_case!(
    dict_literal_unpack_runtime,
    "a = {'x': 1}\nb = {'y': 2}\nc = {**a, **b}\nprint(c['y'])\n",
    "2"
);
runtime_case!(
    dict_literal_unpack_overwrite_runtime,
    "a = {'x': 1}\nb = {'x': 5}\nc = {**a, **b}\nprint(c['x'])\n",
    "5"
);
runtime_case!(
    dict_membership_existing_runtime,
    "d = {'a': 1}\nprint('a' in d)\n",
    "true"
);
runtime_case!(
    dict_membership_missing_runtime,
    "d = {'a': 1}\nprint('z' in d)\n",
    "false"
);
runtime_case!(
    dict_len_after_overwrite_runtime,
    "d = {'a': 1}\nd['a'] = 2\nprint(len(d))\n",
    "1"
);
runtime_case!(
    dict_del_key_runtime,
    "d = {'a': 1, 'b': 2}\ndel d['a']\nprint(len(d))\n",
    "1"
);
runtime_case!(
    dict_nested_lookup_runtime,
    "d = {'outer': {'inner': 42}}\nprint(d['outer']['inner'])\n",
    "42"
);

compile_case!(
    dict_items_loop_compile,
    "d = {'a': 1, 'b': 2}\nfor key, value in d.items():\n    print(key, value)\n"
);
compile_case!(
    dict_keys_loop_compile,
    "d = {'a': 1, 'b': 2}\nfor key in d.keys():\n    print(key)\n"
);
compile_case!(
    dict_values_loop_compile,
    "d = {'a': 1, 'b': 2}\nfor value in d.values():\n    print(value)\n"
);
compile_case!(dict_popitem_compile, "d = {'a': 1}\nitem = d.popitem()\n");
compile_case!(
    dict_update_keyword_args_compile,
    "d = {'a': 1}\nd.update(b=2, c=3)\n"
);
compile_case!(dict_fromkeys_compile, "d = dict.fromkeys(['a', 'b'], 0)\n");
compile_case!(
    dict_comprehension_guard_compile,
    "d = {k: v for k, v in [('a', 1), ('b', 2)] if v > 1}\n"
);
compile_case!(
    dict_comprehension_transform_compile,
    "d = {k.upper(): v * 10 for k, v in [('a', 1), ('b', 2)]}\n"
);
compile_case!(
    dict_setdefault_list_compile,
    "d = {}\nd.setdefault('items', []).append(1)\n"
);
compile_case!(
    dict_union_operator_compile,
    "a = {'x': 1}\nb = {'y': 2}\nc = a | b\n"
);
