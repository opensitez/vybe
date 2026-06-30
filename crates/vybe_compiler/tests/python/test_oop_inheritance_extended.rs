//! OOP: MRO, super(), descriptors, __slots__, class/static vars, metaclass basics.

use crate::helpers::*;

crate::runtime_case!(
    class_instance_attr,
    "class C:\n def __init__(self, x):\n  self.x = x\nprint(C(5).x)\n",
    "5"
);
crate::runtime_case!(
    class_method_calls_instance,
    "class C:\n def f(self):\n  return 1\nprint(C().f())\n",
    "1"
);
crate::runtime_case!(
    inheritance_override,
    "class B:\n def f(self):\n  return 'b'\nclass D(B):\n def f(self):\n  return 'd'\nprint(D().f())\n",
    "d"
);
crate::runtime_case!(
    inheritance_super_call,
    "class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return super().f() + 1\nprint(D().f())\n",
    "2"
);
crate::runtime_case!(
    inheritance_init_chain,
    "class B:\n def __init__(self):\n  self.v = 1\nclass D(B):\n def __init__(self):\n  super().__init__()\n  self.v += 1\nprint(D().v)\n",
    "2"
);
crate::runtime_case!(
    class_variable_shared,
    "class C:\n x = []\nc1 = C()\nc2 = C()\nc1.x.append(1)\nprint(len(c2.x))\n",
    "1"
);
crate::runtime_case!(
    instance_dict_shadows_class,
    "class C:\n x = 1\nc = C()\nc.x = 2\nprint(C.x, c.x)\n",
    "1 2"
);
crate::runtime_case!(
    isinstance_check,
    "class D(Exception):\n pass\nprint(isinstance(D(), Exception))\n",
    "True"
);
crate::runtime_case!(
    issubclass_check,
    "class B: pass\nclass D(B): pass\nprint(issubclass(D, B))\n",
    "True"
);
crate::runtime_case!(
    dunder_repr,
    "class C:\n def __repr__(self):\n  return 'C()'\nprint(repr(C()))\n",
    "C()"
);
crate::runtime_case!(
    dunder_str,
    "class C:\n def __str__(self):\n  return 's'\nprint(str(C()))\n",
    "s"
);
crate::runtime_case!(
    dunder_len,
    "class C:\n def __len__(self):\n  return 3\nprint(len(C()))\n",
    "3"
);
crate::runtime_case!(
    dunder_bool,
    "class C:\n def __bool__(self):\n  return False\nprint(bool(C()))\n",
    "False"
);
crate::runtime_case!(
    dunder_eq,
    "class C:\n def __init__(self, v):\n  self.v = v\n def __eq__(self, o):\n  return self.v == o.v\nprint(C(1) == C(1))\n",
    "True"
);
crate::runtime_case!(
    dunder_lt,
    "class C:\n def __init__(self, v):\n  self.v = v\n def __lt__(self, o):\n  return self.v < o.v\nprint(C(1) < C(2))\n",
    "True"
);
crate::runtime_case!(
    dunder_getitem,
    "class C:\n def __getitem__(self, i):\n  return i * 2\nprint(C()[3])\n",
    "6"
);
crate::runtime_case!(
    dunder_setitem,
    "class C:\n def __init__(self):\n  self.d = {}\n def __setitem__(self, k, v):\n  self.d[k] = v\n def __getitem__(self, k):\n  return self.d[k]\nc = C()\nc['a'] = 1\nprint(c['a'])\n",
    "1"
);
crate::runtime_case!(
    dunder_call,
    "class C:\n def __call__(self, x):\n  return x + 1\nprint(C()(4))\n",
    "5"
);
crate::runtime_case!(
    dunder_iter,
    "class C:\n def __iter__(self):\n  return iter([1, 2])\nprint(list(C()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    dunder_contains,
    "class C:\n def __contains__(self, x):\n  return x == 1\nprint(1 in C())\n",
    "True"
);
crate::runtime_case!(
    property_builtin,
    "class C:\n @property\n def x(self):\n  return 9\nprint(C().x)\n",
    "9"
);
crate::runtime_case!(
    classmethod_factory,
    "class C:\n @classmethod\n def make(cls):\n  return cls()\nprint(isinstance(C.make(), C))\n",
    "True"
);
crate::runtime_case!(
    staticmethod_no_self,
    "class C:\n @staticmethod\n def add(a, b):\n  return a + b\nprint(C.add(2, 3))\n",
    "5"
);
crate::runtime_case!(
    slots_restrict_attrs,
    "class C:\n __slots__ = ('x',)\n def __init__(self, x):\n  self.x = x\nprint(C(1).x)\n",
    "1"
);
crate::runtime_case!(
    mro_diamond_left,
    "class A:\n def f(self):\n  return 'A'\nclass B(A):\n def f(self):\n  return 'B' + super().f()\nclass C(A):\n def f(self):\n  return 'C' + super().f()\nclass D(B, C):\n def f(self):\n  return 'D' + super().f()\nprint(D().f())\n",
    "DBCA"
);
crate::runtime_case!(
    mro_method_resolution,
    "class A:\n def who(self):\n  return 'A'\nclass B(A):\n pass\nclass C(A):\n def who(self):\n  return 'C'\nclass D(B, C):\n pass\nprint(D().who())\n",
    "C"
);
crate::runtime_case!(
  super_two_arg_form,
    "class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return super(D, self).f() + 10\nprint(D().f())\n",
    "11"
);
crate::runtime_case!(
    object_init_default,
    "class C: pass\nprint(isinstance(C(), object))\n",
    "True"
);
crate::runtime_case!(
    type_name,
    "class C: pass\nprint(C.__name__)\n",
    "C"
);
crate::runtime_case!(
    type_bases,
    "class B: pass\nclass D(B): pass\nprint(D.__bases__[0].__name__)\n",
    "B"
);
crate::runtime_case!(
    instance_class_link,
    "class C: pass\nprint(C().__class__ is C)\n",
    "True"
);
crate::runtime_case!(
    class_attr_assignment,
    "class C:\n x = 1\nC.x = 2\nprint(C.x)\n",
    "2"
);
crate::runtime_case!(
    private_name_mangling,
    "class C:\n def __init__(self):\n  self.__x = 1\n def get(self):\n  return self.__x\nprint(C().get())\n",
    "1"
);
crate::runtime_case!(
    dunder_add,
    "class C:\n def __init__(self, v):\n  self.v = v\n def __add__(self, o):\n  return C(self.v + o.v)\nprint((C(1) + C(2)).v)\n",
    "3"
);
crate::runtime_case!(
    dunder_radd,
    "class C:\n def __init__(self, v):\n  self.v = v\n def __radd__(self, o):\n  return C(o + self.v)\nprint((1 + C(2)).v)\n",
    "3"
);
crate::runtime_case!(
    dunder_neg,
    "class C:\n def __init__(self, v):\n  self.v = v\n def __neg__(self):\n  return C(-self.v)\nprint((-C(3)).v)\n",
    "-3"
);
crate::runtime_case!(
    dunder_hash_none_by_default,
    "class C:\n pass\ntry:\n hash(C())\n print('h')\nexcept TypeError:\n print('unhashable')\n",
    "unhashable"
);
crate::runtime_case!(
    dunder_hash_defined,
    "class C:\n def __init__(self, v):\n  self.v = v\n def __hash__(self):\n  return hash(self.v)\nprint(hash(C(1)) == hash(C(1)))\n",
    "True"
);
crate::runtime_case!(
    dataclass_like_manual,
    "class P:\n def __init__(self, x, y):\n  self.x = x\n  self.y = y\nprint(P(1, 2).y)\n",
    "2"
);
crate::runtime_case!(
    nested_class,
    "class O:\n class I:\n  x = 1\nprint(O.I.x)\n",
    "1"
);
crate::runtime_case!(
    class_body_exec_order,
    "class C:\n a = 1\n b = a + 1\nprint(C.b)\n",
    "2"
);
crate::runtime_case!(
    getattr_default,
    "class C:\n x = 1\nprint(getattr(C, 'x'))\n",
    "1"
);
crate::runtime_case!(
    setattr_instance,
    "class C: pass\nc = C()\nsetattr(c, 'y', 3)\nprint(c.y)\n",
    "3"
);
crate::runtime_case!(
    hasattr_false,
    "class C: pass\nprint(hasattr(C(), 'z'))\n",
    "False"
);
crate::runtime_case!(
    delattr_instance,
    "class C:\n def __init__(self):\n  self.x = 1\nc = C()\ndelattr(c, 'x')\nprint(hasattr(c, 'x'))\n",
    "False"
);
crate::runtime_case!(
    dir_instance,
    "class C:\n x = 1\nc = C()\nprint('x' in dir(c))\n",
    "True"
);
crate::runtime_case!(
    vars_instance,
    "class C:\n def __init__(self):\n  self.a = 1\nprint(vars(C())['a'])\n",
    "1"
);

crate::compile_case!(oop_metaclass_type, "class M(type):\n pass\nclass C(metaclass=M):\n pass\n");
crate::compile_case!(oop_abstract_base, "from abc import ABC, abstractmethod\nclass B(ABC):\n @abstractmethod\n def m(self):\n  pass\n");
crate::compile_case!(oop_slots_subclass, "class B:\n __slots__ = ()\nclass D(B):\n __slots__ = ('x',)\n");
crate::compile_case!(oop_descriptors, "class D:\n def __get__(self, obj, owner):\n  return 1\nclass C:\n x = D()\n");
crate::compile_case!(oop_init_subclass, "class B:\n def __init_subclass__(cls, **kw):\n  pass\nclass D(B):\n pass\n");
