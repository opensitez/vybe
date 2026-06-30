//! Extended dict methods: get/setdefault, popitem, merge, views, comprehension patterns.


crate::runtime_case!(
    dict_get_existing,
    "d = {'a': 1}\nprint(d.get('a'))\n",
    "1"
);
crate::runtime_case!(
    dict_get_missing_default,
    "d = {}\nprint(d.get('x', 99))\n",
    "99"
);
crate::runtime_case!(
    dict_get_missing_none,
    "d = {}\nprint(d.get('x'))\n",
    "None"
);
crate::runtime_case!(
    dict_setdefault_inserts,
    "d = {}\nd.setdefault('k', 5)\nprint(d['k'])\n",
    "5"
);
crate::runtime_case!(
    dict_setdefault_keeps_existing,
    "d = {'k': 1}\nd.setdefault('k', 9)\nprint(d['k'])\n",
    "1"
);
crate::runtime_case!(
    dict_pop_existing,
    "d = {'a': 1, 'b': 2}\nprint(d.pop('a'))\n",
    "1"
);
crate::runtime_case!(
    dict_pop_missing_default,
    "d = {}\nprint(d.pop('x', 0))\n",
    "0"
);
crate::runtime_case!(
    dict_popitem_lifo,
    "d = {'a': 1, 'b': 2}\nk, v = d.popitem()\nprint(k, v)\n",
    "b 2"
);
crate::runtime_case!(
    dict_update_mapping,
    "d = {'a': 1}\nd.update({'b': 2})\nprint(sorted(d.items()))\n",
    "[('a', 1), ('b', 2)]"
);
crate::runtime_case!(
    dict_update_kwargs,
    "d = {'a': 1}\nd.update(c=3)\nprint(d['c'])\n",
    "3"
);
crate::runtime_case!(
    dict_keys_membership,
    "d = {'x': 1}\nprint('x' in d)\n",
    "True"
);
crate::runtime_case!(
    dict_values_contains,
    "d = {'a': 10, 'b': 20}\nprint(10 in d.values())\n",
    "True"
);
crate::runtime_case!(
    dict_items_len,
    "d = {'a': 1, 'b': 2, 'c': 3}\nprint(len(d.items()))\n",
    "3"
);
crate::runtime_case!(
    dict_clear_then_len,
    "d = {'a': 1}\nd.clear()\nprint(len(d))\n",
    "0"
);
crate::runtime_case!(
    dict_copy_independent,
    "d = {'a': 1}\ne = d.copy()\ne['b'] = 2\nprint(len(d), len(e))\n",
    "1 2"
);
crate::runtime_case!(
    dict_comp_from_range,
    "print({x: x * x for x in range(4)})\n",
    "{0: 0, 1: 1, 2: 4, 3: 9}"
);
crate::runtime_case!(
    dict_comp_filter_even,
    "print({x: x for x in range(5) if x % 2 == 0})\n",
    "{0: 0, 2: 2, 4: 4}"
);
crate::runtime_case!(
    dict_comp_swap_kv,
    "print({v: k for k, v in [('a', 1), ('b', 2)]})\n",
    "{1: 'a', 2: 'b'}"
);
crate::runtime_case!(
    dict_comp_nested_len,
    "print({k: len(k) for k in ['aa', 'bbb']})\n",
    "{'aa': 2, 'bbb': 3}"
);
crate::runtime_case!(
    dict_merge_pipe_operator,
    "a = {'x': 1}\nb = {'y': 2}\nc = a | b\nprint(sorted(c.keys()))\n",
    "['x', 'y']"
);
crate::runtime_case!(
    dict_merge_inplace_pipe,
    "a = {'x': 1}\na |= {'y': 2}\nprint('y' in a)\n",
    "True"
);
crate::runtime_case!(
    dict_fromkeys_default,
    "d = dict.fromkeys(['a', 'b'], 0)\nprint(d['b'])\n",
    "0"
);
crate::runtime_case!(
    dict_fromkeys_none_default,
    "d = dict.fromkeys('xy')\nprint(d['x'])\n",
    "None"
);
crate::runtime_case!(
    dict_literal_nested_access,
    "d = {'outer': {'inner': 9}}\nprint(d['outer']['inner'])\n",
    "9"
);
crate::runtime_case!(
    dict_del_key,
    "d = {'a': 1, 'b': 2}\ndel d['a']\nprint(list(d))\n",
    "['b']"
);
crate::runtime_case!(
    dict_bool_nonempty,
    "print(bool({'a': 1}))\n",
    "True"
);
crate::runtime_case!(
    dict_bool_empty,
    "print(bool({}))\n",
    "False"
);
crate::runtime_case!(
    dict_equality_same_pairs,
    "print({'a': 1, 'b': 2} == {'b': 2, 'a': 1})\n",
    "True"
);
crate::runtime_case!(
    dict_inequality_extra_key,
    "print({'a': 1} == {'a': 1, 'b': 2})\n",
    "False"
);
crate::runtime_case!(
    dict_keys_iterable_sorted,
    "print(sorted({'c': 3, 'a': 1, 'b': 2}))\n",
    "['a', 'b', 'c']"
);
crate::runtime_case!(
    dict_values_list_sum,
    "print(sum({'a': 1, 'b': 2, 'c': 3}.values()))\n",
    "6"
);
crate::runtime_case!(
    dict_items_unpack_loop,
    "d = {'x': 10, 'y': 20}\ns = 0\nfor k, v in d.items():\n    s += v\nprint(s)\n",
    "30"
);
crate::runtime_case!(
    dict_get_chain_default,
    "d = {'a': {'b': 1}}\nprint(d.get('a', {}).get('b', 0))\n",
    "1"
);
crate::runtime_case!(
    dict_comp_if_else_value,
    "print({x: ('even' if x % 2 == 0 else 'odd') for x in range(3)})\n",
    "{0: 'even', 1: 'odd', 2: 'even'}"
);
crate::runtime_case!(
    dict_union_overwrite,
    "a = {'k': 1}\nb = {'k': 2}\nprint((a | b)['k'])\n",
    "2"
);
crate::runtime_case!(
    dict_reversed_keys_order,
    "d = {'a': 1, 'b': 2, 'c': 3}\nprint(list(reversed(d)))\n",
    "['c', 'b', 'a']"
);
crate::runtime_case!(
    dict_popitem_empty_raises_compile,
    "d = {}\ntry:\n    d.popitem()\n    print('ok')\nexcept KeyError:\n    print('err')\n",
    "err"
);
crate::runtime_case!(
    dict_setitem_new_key,
    "d = {}\nd['n'] = 7\nprint(d['n'])\n",
    "7"
);
crate::runtime_case!(
    dict_setitem_overwrite,
    "d = {'k': 1}\nd['k'] = 9\nprint(d['k'])\n",
    "9"
);
crate::runtime_case!(
    dict_unpack_in_literal,
    "base = {'a': 1}\nmerged = {**base, 'b': 2}\nprint(merged['b'])\n",
    "2"
);
crate::runtime_case!(
    dict_enumerate_items,
    "d = {'x': 1, 'y': 2}\nprint([i for i, _ in enumerate(d)])\n",
    "[0, 1]"
);
crate::runtime_case!(
    dict_zip_keys_values,
    "d = {'a': 1, 'b': 2}\nprint(list(zip(d, d.values())))\n",
    "[('a', 1), ('b', 2)]"
);
crate::runtime_case!(
    dict_any_value_truthy,
    "print(any({'a': 0, 'b': 1, 'c': 0}.values()))\n",
    "True"
);
crate::runtime_case!(
    dict_all_values_truthy,
    "print(all({'a': 1, 'b': 2}.values()))\n",
    "True"
);

crate::compile_case!(dict_popitem_twice, "d = {'a': 1, 'b': 2}\nd.popitem()\nd.popitem()\n");
crate::compile_case!(dict_update_iterable, "d = {}\nd.update([('a', 1), ('b', 2)])\n");
crate::compile_case!(dict_view_set_difference, "d = {'a': 1, 'b': 2}\nk = d.keys()\n");
crate::compile_case!(dict_comp_nested, "d = {k: {k: k} for k in range(2)}\n");
crate::compile_case!(dict_del_missing_compile, "d = {}\ntry:\n    del d['x']\nexcept KeyError:\n    pass\n");
