use crate::helpers::{run_print, run_python_one};

#[test]
fn ordered_dict_preserves_insertion_order_keys() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nd = OrderedDict([('b', 2), ('a', 1), ('c', 3)])\nprint(list(d.keys()))\n"
        ),
        "['b', 'a', 'c']"
    );
}

#[test]
fn ordered_dict_move_to_end() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nd = OrderedDict(a=1, b=2)\nd.move_to_end('a')\nprint(list(d.keys()))\n"
        ),
        "['b', 'a']"
    );
}

#[test]
fn ordered_dict_popitem_last() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nd = OrderedDict(x=1, y=2)\nprint(d.popitem())\n"
        ),
        "('y', 2)"
    );
}

#[test]
fn ordered_dict_popitem_first() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nd = OrderedDict(x=1, y=2)\nprint(d.popitem(last=False))\n"
        ),
        "('x', 1)"
    );
}

#[test]
fn ordered_dict_equality_order_sensitive() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\na = OrderedDict([('a', 1), ('b', 2)])\nb = OrderedDict([('b', 2), ('a', 1)])\nprint(a == b)\n"
        ),
        "False"
    );
}

#[test]
fn counter_from_list_counts() {
    assert_eq!(
        run_python_one("from collections import Counter\nprint(list(Counter('aabbc').items()))\n"),
        "[('a', 2), ('b', 2), ('c', 1)]"
    );
}

#[test]
fn counter_most_common() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter([1, 1, 2, 3, 3, 3])\nprint(c.most_common(1)[0])\n"
        ),
        "(3, 3)"
    );
}

#[test]
fn counter_add_counts() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter(a=3, b=1)\nc.update(Counter(a=1, b=2))\nprint(c['a'], c['b'])\n"
        ),
        "4 3"
    );
}

#[test]
fn counter_subtract() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter(a=5, b=2)\nc.subtract(Counter(a=2, b=2))\nprint(c['a'], c['b'])\n"
        ),
        "3 0"
    );
}

#[test]
fn counter_elements_positive_only() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter(a=2, b=0)\nprint(sorted(c.elements()))\n"
        ),
        "['a', 'a']"
    );
}

#[test]
fn counter_total() {
    assert_eq!(
        run_python_one("from collections import Counter\nprint(sum(Counter('abba').values()))\n"),
        "4"
    );
}

#[test]
fn deque_append_and_pop() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([1, 2])\nd.append(3)\nprint(d.pop())\n"
        ),
        "3"
    );
}

#[test]
fn deque_appendleft_and_popleft() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([2, 3])\nd.appendleft(1)\nprint(d.popleft())\n"
        ),
        "1"
    );
}

#[test]
fn deque_rotate_right() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([1, 2, 3])\nd.rotate(1)\nprint(list(d))\n"
        ),
        "[3, 1, 2]"
    );
}

#[test]
fn deque_rotate_left() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([1, 2, 3])\nd.rotate(-1)\nprint(list(d))\n"
        ),
        "[2, 3, 1]"
    );
}

#[test]
fn deque_maxlen_drops_left() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque(maxlen=2)\nd.extend([1, 2, 3])\nprint(list(d))\n"
        ),
        "[2, 3]"
    );
}

#[test]
fn defaultdict_int_factory() {
    assert_eq!(
        run_python_one(
            "from collections import defaultdict\nd = defaultdict(int)\nd['a'] += 1\nd['a'] += 1\nprint(d['a'])\n"
        ),
        "2"
    );
}

#[test]
fn defaultdict_list_factory() {
    assert_eq!(
        run_python_one(
            "from collections import defaultdict\nd = defaultdict(list)\nd['k'].append(1)\nd['k'].append(2)\nprint(d['k'])\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn defaultdict_missing_key_returns_default() {
    assert_eq!(
        run_python_one(
            "from collections import defaultdict\nd = defaultdict(lambda: 9)\nprint(d['missing'])\n"
        ),
        "9"
    );
}

#[test]
fn namedtuple_field_access() {
    assert_eq!(
        run_python_one(
            "from collections import namedtuple\nPoint = namedtuple('Point', 'x y')\np = Point(1, 2)\nprint(p.x, p.y)\n"
        ),
        "1 2"
    );
}

#[test]
fn namedtuple_unpack() {
    assert_eq!(
        run_python_one(
            "from collections import namedtuple\nPair = namedtuple('Pair', 'a b')\na, b = Pair(3, 4)\nprint(a + b)\n"
        ),
        "7"
    );
}

#[test]
fn namedtuple_asdict() {
    assert_eq!(
        run_python_one(
            "from collections import namedtuple\nR = namedtuple('R', 'x y')\nprint(R(1, 2)._asdict())\n"
        ),
        "{'x': 1, 'y': 2}"
    );
}

#[test]
fn chainmap_get_first_mapping() {
    assert_eq!(
        run_python_one(
            "from collections import ChainMap\ncm = ChainMap({'a': 1}, {'a': 2, 'b': 3})\nprint(cm['a'], cm['b'])\n"
        ),
        "1 3"
    );
}

#[test]
fn chainmap_new_child() {
    assert_eq!(
        run_python_one(
            "from collections import ChainMap\nbase = ChainMap({'x': 1})\nchild = base.new_child({'x': 2})\nprint(child['x'], base['x'])\n"
        ),
        "2 1"
    );
}

#[test]
fn userdict_overrides_getitem() {
    assert_eq!(
        run_python_one(
            "from collections import UserDict\nclass D(UserDict):\n pass\nd = D({'a': 1})\nprint(d['a'])\n"
        ),
        "1"
    );
}

#[test]
fn userlist_append() {
    assert_eq!(
        run_python_one(
            "from collections import UserList\nul = UserList([1])\nul.append(2)\nprint(ul)\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn userstring_upper() {
    assert_eq!(
        run_python_one(
            "from collections import UserString\ns = UserString('ab')\nprint(str(s).upper())\n"
        ),
        "AB"
    );
}

#[test]
fn counter_from_keys() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter.fromkeys(['a', 'b', 'a'], v=2)\nprint(c['a'], c['b'])\n"
        ),
        "2 2"
    );
}

#[test]
fn deque_extend_left() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([3])\nd.extendleft([2, 1])\nprint(list(d))\n"
        ),
        "[1, 2, 3]"
    );
}

#[test]
fn deque_clear() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([1, 2])\nd.clear()\nprint(len(d))\n"
        ),
        "0"
    );
}

#[test]
fn ordered_dict_reinsert_moves_position() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nd = OrderedDict(a=1, b=2)\nd['a'] = 9\nprint(list(d.keys()))\n"
        ),
        "['a', 'b']"
    );
}

#[test]
fn counter_bit_and_intersection() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter(a=3, b=1) & Counter(a=1, c=2)\nprint(dict(c))\n"
        ),
        "{'a': 1}"
    );
}

#[test]
fn counter_bit_or_union() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter(a=3, b=1) | Counter(a=1, c=2)\nprint(sorted(c.items()))\n"
        ),
        "[('a', 3), ('b', 1), ('c', 2)]"
    );
}

#[test]
fn defaultdict_set_factory() {
    assert_eq!(
        run_python_one(
            "from collections import defaultdict\nd = defaultdict(set)\nd['g'].add(1)\nd['g'].add(2)\nprint(sorted(d['g']))\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn namedtuple_replace() {
    assert_eq!(
        run_python_one(
            "from collections import namedtuple\nP = namedtuple('P', 'x y')\nprint(P(1, 2)._replace(y=9))\n"
        ),
        "P(x=1, y=9)"
    );
}

#[test]
fn deque_index_method() {
    assert_eq!(
        run_python_one(
            "from collections import deque\nd = deque([10, 20, 30])\nprint(d.index(20))\n"
        ),
        "1"
    );
}

#[test]
fn deque_count_method() {
    assert_eq!(
        run_python_one("from collections import deque\nd = deque([1, 1, 2])\nprint(d.count(1))\n"),
        "2"
    );
}

#[test]
fn ordered_dict_len() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nprint(len(OrderedDict([('a', 1), ('b', 2)])))\n"
        ),
        "2"
    );
}

#[test]
fn counter_len_unique_keys() {
    assert_eq!(
        run_python_one("from collections import Counter\nprint(len(Counter('abracadabra')))\n"),
        "5"
    );
}

#[test]
fn chainmap_maps_property() {
    assert_eq!(
        run_python_one(
            "from collections import ChainMap\ncm = ChainMap({'a': 1}, {'b': 2})\nprint(len(cm.maps))\n"
        ),
        "2"
    );
}

#[test]
fn deque_iter_protocol() {
    assert_eq!(
        run_python_one("from collections import deque\nprint(sum(deque([1, 2, 3])))\n"),
        "6"
    );
}

#[test]
fn defaultdict_iter_keys() {
    assert_eq!(
        run_python_one(
            "from collections import defaultdict\nd = defaultdict(int, {'x': 1})\nprint(sorted(d))\n"
        ),
        "['x']"
    );
}

#[test]
fn namedtuple_fields_attribute() {
    assert_eq!(
        run_python_one(
            "from collections import namedtuple\nT = namedtuple('T', 'a b')\nprint(T._fields)\n"
        ),
        "('a', 'b')"
    );
}

#[test]
fn counter_negative_counts_allowed() {
    assert_eq!(
        run_python_one(
            "from collections import Counter\nc = Counter(a=1)\nc.subtract(Counter(a=3))\nprint(c['a'])\n"
        ),
        "-2"
    );
}

#[test]
fn ordered_dict_get_method() {
    assert_eq!(
        run_python_one(
            "from collections import OrderedDict\nd = OrderedDict(a=1)\nprint(d.get('b', 0))\n"
        ),
        "0"
    );
}
