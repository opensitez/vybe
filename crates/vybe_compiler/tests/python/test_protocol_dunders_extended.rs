//! Operator dunders runtime: __add__, __mul__, comparisons, container protocol.

crate::runtime_case!(
    dunder_add,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __add__(self, o):\n  return V(self.v + o.v)\nprint((V(1) + V(2)).v)\n",
    "3"
);
crate::runtime_case!(
    dunder_radd,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __radd__(self, o):\n  return V(o + self.v)\nprint((1 + V(2)).v)\n",
    "3"
);
crate::runtime_case!(
    dunder_sub,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __sub__(self, o):\n  return V(self.v - o.v)\nprint((V(5) - V(2)).v)\n",
    "3"
);
crate::runtime_case!(
    dunder_mul,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __mul__(self, o):\n  return V(self.v * o.v)\nprint((V(3) * V(4)).v)\n",
    "12"
);
crate::runtime_case!(
    dunder_truediv,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __truediv__(self, o):\n  return V(self.v / o.v)\nprint((V(10) / V(2)).v)\n",
    "5.0"
);
crate::runtime_case!(
    dunder_floordiv,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __floordiv__(self, o):\n  return V(self.v // o.v)\nprint((V(7) // V(2)).v)\n",
    "3"
);
crate::runtime_case!(
    dunder_mod,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __mod__(self, o):\n  return V(self.v % o.v)\nprint((V(10) % V(3)).v)\n",
    "1"
);
crate::runtime_case!(
    dunder_pow,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __pow__(self, o):\n  return V(self.v ** o.v)\nprint((V(2) ** V(3)).v)\n",
    "8"
);
crate::runtime_case!(
    dunder_neg,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __neg__(self):\n  return V(-self.v)\nprint((-V(3)).v)\n",
    "-3"
);
crate::runtime_case!(
    dunder_pos,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __pos__(self):\n  return V(+self.v)\nprint((+V(-3)).v)\n",
    "-3"
);
crate::runtime_case!(
    dunder_abs,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __abs__(self):\n  return V(abs(self.v))\nprint(abs(V(-5)).v)\n",
    "5"
);
crate::runtime_case!(
    dunder_eq,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __eq__(self, o):\n  return self.v == o.v\nprint(V(1) == V(1))\n",
    "True"
);
crate::runtime_case!(
    dunder_lt,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __lt__(self, o):\n  return self.v < o.v\nprint(V(1) < V(2))\n",
    "True"
);
crate::runtime_case!(
    dunder_le,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __le__(self, o):\n  return self.v <= o.v\nprint(V(2) <= V(2))\n",
    "True"
);
crate::runtime_case!(
    dunder_gt,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __gt__(self, o):\n  return self.v > o.v\nprint(V(3) > V(2))\n",
    "True"
);
crate::runtime_case!(
    dunder_ge,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __ge__(self, o):\n  return self.v >= o.v\nprint(V(2) >= V(1))\n",
    "True"
);
crate::runtime_case!(
    dunder_ne,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __ne__(self, o):\n  return self.v != o.v\nprint(V(1) != V(2))\n",
    "True"
);
crate::runtime_case!(
    dunder_len,
    "class V:\n def __len__(self):\n  return 3\nprint(len(V()))\n",
    "3"
);
crate::runtime_case!(
    dunder_getitem,
    "class V:\n def __getitem__(self, i):\n  return i * 2\nprint(V()[3])\n",
    "6"
);
crate::runtime_case!(
    dunder_setitem,
    "class V:\n def __init__(self):\n  self.d = {}\n def __setitem__(self, k, v):\n  self.d[k] = v\n def __getitem__(self, k):\n  return self.d[k]\nv = V()\nv['a'] = 1\nprint(v['a'])\n",
    "1"
);
crate::runtime_case!(
    dunder_delitem,
    "class V:\n def __init__(self):\n  self.d = {'a': 1}\n def __delitem__(self, k):\n  del self.d[k]\n def __getitem__(self, k):\n  return self.d[k]\nv = V()\ndel v['a']\nprint('a' in v.d)\n",
    "False"
);
crate::runtime_case!(
    dunder_contains,
    "class V:\n def __contains__(self, x):\n  return x == 1\nprint(1 in V())\n",
    "True"
);
crate::runtime_case!(
    dunder_iter,
    "class V:\n def __iter__(self):\n  return iter([1, 2])\nprint(list(V()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    dunder_next,
    "class V:\n def __init__(self):\n  self.n = 0\n def __iter__(self):\n  return self\n def __next__(self):\n  if self.n >= 2:\n   raise StopIteration\n  self.n += 1\n  return self.n\nprint(list(V()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    dunder_call,
    "class V:\n def __call__(self, x):\n  return x + 1\nprint(V()(4))\n",
    "5"
);
crate::runtime_case!(
    dunder_str,
    "class V:\n def __str__(self):\n  return 'str'\nprint(str(V()))\n",
    "str"
);
crate::runtime_case!(
    dunder_repr,
    "class V:\n def __repr__(self):\n  return 'V()'\nprint(repr(V()))\n",
    "V()"
);
crate::runtime_case!(
    dunder_bool,
    "class V:\n def __bool__(self):\n  return False\nprint(bool(V()))\n",
    "False"
);
crate::runtime_case!(
    dunder_hash,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __hash__(self):\n  return hash(self.v)\nprint(hash(V(1)) == hash(V(1)))\n",
    "True"
);
crate::runtime_case!(
    dunder_int,
    "class V:\n def __int__(self):\n  return 7\nprint(int(V()))\n",
    "7"
);
crate::runtime_case!(
    dunder_float,
    "class V:\n def __float__(self):\n  return 3.5\nprint(float(V()))\n",
    "3.5"
);
crate::runtime_case!(
    dunder_index,
    "class V:\n def __index__(self):\n  return 2\nprint([0, 1, 2, 3][V()])\n",
    "2"
);
crate::runtime_case!(
    dunder_enter_exit,
    "class V:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nwith V():\n print('in')\n",
    "in"
);
crate::runtime_case!(
    dunder_iadd,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __iadd__(self, o):\n  self.v += o.v\n  return self\na = V(1)\na += V(2)\nprint(a.v)\n",
    "3"
);
crate::runtime_case!(
    dunder_imul,
    "class V:\n def __init__(self, v):\n  self.v = v\n def __imul__(self, o):\n  self.v *= o.v\n  return self\na = V(2)\na *= V(3)\nprint(a.v)\n",
    "6"
);
crate::runtime_case!(
    dunder_matmul,
    "class V:\n def __matmul__(self, o):\n  return 'mat'\nprint(V() @ V())\n",
    "mat"
);
crate::runtime_case!(
    dunder_and,
    "class V:\n def __and__(self, o):\n  return 'and'\nprint(V() & V())\n",
    "and"
);
crate::runtime_case!(
    dunder_or,
    "class V:\n def __or__(self, o):\n  return 'or'\nprint(V() | V())\n",
    "or"
);
crate::runtime_case!(
    dunder_xor,
    "class V:\n def __xor__(self, o):\n  return 'xor'\nprint(V() ^ V())\n",
    "xor"
);
crate::runtime_case!(
    dunder_lshift,
    "class V:\n def __lshift__(self, o):\n  return 'ls'\nprint(V() << V())\n",
    "ls"
);
crate::runtime_case!(
    dunder_rshift,
    "class V:\n def __rshift__(self, o):\n  return 'rs'\nprint(V() >> V())\n",
    "rs"
);
crate::runtime_case!(
    dunder_invert,
    "class V:\n def __invert__(self):\n  return 'inv'\nprint(~V())\n",
    "inv"
);
crate::runtime_case!(
    dunder_format,
    "class V:\n def __format__(self, spec):\n  return 'fmt'\nprint(f'{V():x}')\n",
    "fmt"
);
crate::runtime_case!(
    dunder_bytes,
    "class V:\n def __bytes__(self):\n  return b'v'\nprint(bytes(V()))\n",
    "b'v'"
);
crate::runtime_case!(
    dunder_reversed,
    "class V:\n def __reversed__(self):\n  return iter([3, 2, 1])\nprint(list(reversed(V())))\n",
    "[3, 2, 1]"
);
crate::runtime_case!(
    dunder_length_hint,
    "class V:\n def __length_hint__(self):\n  return 5\nprint(V().__length_hint__())\n",
    "5"
);

crate::compile_case!(dunder_await, "class V:\n def __await__(self):\n  yield 1\n");
crate::compile_case!(
    dunder_aiter_aenter,
    "class V:\n async def __aenter__(self): return self\n async def __aexit__(self, *a): pass\n"
);
crate::compile_case!(
    dunder_getnewargs,
    "class V:\n def __getnewargs__(self):\n  return ()\n"
);
crate::compile_case!(
    dunder_reduce,
    "class V:\n def __reduce__(self):\n  return (V, ())\n"
);
crate::compile_case!(
    dunder_copy,
    "class V:\n def __copy__(self):\n  return V()\n"
);
