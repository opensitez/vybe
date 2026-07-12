//! collections.Counter, defaultdict, deque, OrderedDict runtime patterns.

crate::runtime_case!(
    counter_from_list,
    "from collections import Counter\nprint(Counter([1, 1, 2]))\n",
    "Counter({1: 2, 2: 1})"
);
crate::runtime_case!(
    counter_most_common,
    "from collections import Counter\nc = Counter('abracadabra')\nprint(c.most_common(1)[0][0])\n",
    "a"
);
crate::runtime_case!(
    counter_elements,
    "from collections import Counter\nc = Counter(a=2, b=1)\nprint(sorted(c.elements()))\n",
    "['a', 'a', 'b']"
);
crate::runtime_case!(
    counter_subtract,
    "from collections import Counter\nc = Counter(a=3, b=1)\nc.subtract({'a': 1})\nprint(c['a'])\n",
    "2"
);
crate::runtime_case!(
    counter_update,
    "from collections import Counter\nc = Counter(a=1)\nc.update({'a': 2, 'b': 1})\nprint(c['b'])\n",
    "1"
);
crate::runtime_case!(
    counter_total,
    "from collections import Counter\nprint(sum(Counter('aab').values()))\n",
    "3"
);
crate::runtime_case!(
    counter_add_another,
    "from collections import Counter\nprint(Counter(a=1) + Counter(a=2))\n",
    "Counter({'a': 3})"
);
crate::runtime_case!(
    counter_sub_another,
    "from collections import Counter\nprint(Counter(a=3) - Counter(a=1))\n",
    "Counter({'a': 2})"
);
crate::runtime_case!(
    counter_and_intersection,
    "from collections import Counter\nprint(Counter(a=2, b=1) & Counter(a=1, c=1))\n",
    "Counter({'a': 1})"
);
crate::runtime_case!(
    counter_or_union,
    "from collections import Counter\nprint(Counter(a=1) | Counter(b=2))\n",
    "Counter({'a': 1, 'b': 2})"
);
crate::runtime_case!(
    defaultdict_list_default,
    "from collections import defaultdict\nd = defaultdict(list)\nd['k'].append(1)\nprint(d['k'])\n",
    "[1]"
);
crate::runtime_case!(
    defaultdict_int_default,
    "from collections import defaultdict\nd = defaultdict(int)\nd['x'] += 1\nprint(d['x'])\n",
    "1"
);
crate::runtime_case!(
    defaultdict_factory_zero,
    "from collections import defaultdict\nd = defaultdict(lambda: 0)\nprint(d['missing'])\n",
    "0"
);
crate::runtime_case!(
    defaultdict_existing_key,
    "from collections import defaultdict\nd = defaultdict(int)\nd['a'] = 5\nprint(d['a'])\n",
    "5"
);
crate::runtime_case!(
    deque_append,
    "from collections import deque\nd = deque([1])\nd.append(2)\nprint(list(d))\n",
    "[1, 2]"
);
crate::runtime_case!(
    deque_appendleft,
    "from collections import deque\nd = deque([2])\nd.appendleft(1)\nprint(list(d))\n",
    "[1, 2]"
);
crate::runtime_case!(
    deque_pop,
    "from collections import deque\nd = deque([1, 2])\nprint(d.pop())\n",
    "2"
);
crate::runtime_case!(
    deque_popleft,
    "from collections import deque\nd = deque([1, 2])\nprint(d.popleft())\n",
    "1"
);
crate::runtime_case!(
    deque_rotate,
    "from collections import deque\nd = deque([1, 2, 3])\nd.rotate(1)\nprint(list(d))\n",
    "[3, 1, 2]"
);
crate::runtime_case!(
    deque_rotate_negative,
    "from collections import deque\nd = deque([1, 2, 3])\nd.rotate(-1)\nprint(list(d))\n",
    "[2, 3, 1]"
);
crate::runtime_case!(
    deque_maxlen_evict,
    "from collections import deque\nd = deque(maxlen=2)\nd.extend([1, 2, 3])\nprint(list(d))\n",
    "[2, 3]"
);
crate::runtime_case!(
    deque_extend,
    "from collections import deque\nd = deque([1])\nd.extend([2, 3])\nprint(len(d))\n",
    "3"
);
crate::runtime_case!(
    deque_extendleft,
    "from collections import deque\nd = deque([3])\nd.extendleft([2, 1])\nprint(list(d))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    deque_clear,
    "from collections import deque\nd = deque([1, 2])\nd.clear()\nprint(len(d))\n",
    "0"
);
crate::runtime_case!(
    deque_count,
    "from collections import deque\nprint(deque([1, 2, 1]).count(1))\n",
    "2"
);
crate::runtime_case!(
    deque_reverse,
    "from collections import deque\nd = deque([1, 2, 3])\nd.reverse()\nprint(list(d))\n",
    "[3, 2, 1]"
);
crate::runtime_case!(
    namedtuple_basic,
    "from collections import namedtuple\nP = namedtuple('P', 'x y')\np = P(1, 2)\nprint(p.x, p.y)\n",
    "1 2"
);
crate::runtime_case!(
    namedtuple_asdict,
    "from collections import namedtuple\nP = namedtuple('P', 'a b')\nprint(namedtuple('P', 'a b')(1, 2)._asdict())\n",
    "{'a': 1, 'b': 2}"
);
crate::runtime_case!(
    namedtuple_replace,
    "from collections import namedtuple\nP = namedtuple('P', 'a b')\np = P(1, 2)\nprint(p._replace(a=9).a)\n",
    "9"
);
crate::runtime_case!(
    ordereddict_insertion,
    "from collections import OrderedDict\nd = OrderedDict()\nd['b'] = 2\nd['a'] = 1\nprint(list(d.keys()))\n",
    "['b', 'a']"
);
crate::runtime_case!(
    ordereddict_move_to_end,
    "from collections import OrderedDict\nd = OrderedDict(a=1, b=2)\nd.move_to_end('a')\nprint(list(d.keys())[-1])\n",
    "a"
);
crate::runtime_case!(
    chainmap_lookup,
    "from collections import ChainMap\ncm = ChainMap({'a': 1}, {'b': 2})\nprint(cm['a'], cm['b'])\n",
    "1 2"
);
crate::runtime_case!(
    chainmap_shadow,
    "from collections import ChainMap\ncm = ChainMap({'x': 1}, {'x': 9})\nprint(cm['x'])\n",
    "1"
);
crate::runtime_case!(
    userdict_with_default,
    "from collections import UserDict\nclass D(UserDict):\n pass\nd = D({'a': 1})\nprint(d['a'])\n",
    "1"
);
crate::runtime_case!(
    userlist_append,
    "from collections import UserList\nul = UserList([1])\nul.append(2)\nprint(ul)\n",
    "[1, 2]"
);
crate::runtime_case!(
    userstring_upper,
    "from collections import UserString\nprint(UserString('ab').upper())\n",
    "AB"
);
crate::runtime_case!(
    counter_from_string,
    "from collections import Counter\nprint(dict(Counter('hello')))\n",
    "{'h': 1, 'e': 1, 'l': 2, 'o': 1}"
);
crate::runtime_case!(
    deque_index,
    "from collections import deque\nprint(deque([10, 20, 30]).index(20))\n",
    "1"
);
crate::runtime_case!(
    deque_remove,
    "from collections import deque\nd = deque([1, 2, 3])\nd.remove(2)\nprint(list(d))\n",
    "[1, 3]"
);
crate::runtime_case!(
    defaultdict_setdefault_differs,
    "from collections import defaultdict\nd = defaultdict(int)\nprint(d.get('z', 99))\n",
    "99"
);
crate::runtime_case!(
    counter_negative_count,
    "from collections import Counter\nc = Counter(a=1)\nc['a'] -= 2\nprint(c['a'])\n",
    "-1"
);
crate::runtime_case!(
    deque_bool_nonempty,
    "from collections import deque\nprint(bool(deque([1])))\n",
    "True"
);
crate::runtime_case!(
    deque_bool_empty,
    "from collections import deque\nprint(bool(deque()))\n",
    "False"
);
crate::runtime_case!(
    namedtuple_len,
    "from collections import namedtuple\nP = namedtuple('P', 'x y')\nprint(len(P(1, 2)))\n",
    "2"
);
crate::runtime_case!(
    counter_bool,
    "from collections import Counter\nprint(bool(Counter()))\n",
    "False"
);

crate::compile_case!(
    counter_in_place_add,
    "from collections import Counter\nc = Counter(a=1)\nc += Counter(a=2)\n"
);
crate::compile_case!(
    deque_maxlen_none,
    "from collections import deque\nd = deque(maxlen=None)\n"
);
crate::compile_case!(
    ordereddict_popitem,
    "from collections import OrderedDict\nd = OrderedDict(a=1, b=2)\nd.popitem(last=False)\n"
);
crate::compile_case!(
    chainmap_new_child,
    "from collections import ChainMap\ncm = ChainMap({'a': 1})\ncm.new_child({'b': 2})\n"
);
crate::compile_case!(
    namedtuple_defaults,
    "from collections import namedtuple\nP = namedtuple('P', 'x y', defaults=[0])\nP(1)\n"
);
