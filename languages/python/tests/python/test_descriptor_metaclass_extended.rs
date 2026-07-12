//! Descriptor protocol, property, __slots__, metaclass runtime.

crate::runtime_case!(
    descriptor_get,
    "class D:\n def __get__(self, obj, owner):\n  return 42\nclass C:\n x = D()\nprint(C().x)\n",
    "42"
);
crate::runtime_case!(
    descriptor_set,
    "class D:\n def __set__(self, obj, val):\n  obj._v = val\nclass C:\n x = D()\nc = C()\nc.x = 9\nprint(c._v)\n",
    "9"
);
crate::runtime_case!(
    descriptor_delete,
    "class D:\n def __delete__(self, obj):\n  obj._v = None\nclass C:\n x = D()\n def __init__(self):\n  self._v = 1\nc = C()\ndel c.x\nprint(c._v is None)\n",
    "True"
);
crate::runtime_case!(
    property_getter,
    "class C:\n @property\n def x(self):\n  return 1\nprint(C().x)\n",
    "1"
);
crate::runtime_case!(
    property_setter,
    "class C:\n @property\n def x(self):\n  return self._x\n @x.setter\n def x(self, v):\n  self._x = v\nc = C()\nc.x = 5\nprint(c.x)\n",
    "5"
);
crate::runtime_case!(
    property_deleter,
    "class C:\n @property\n def x(self):\n  return self._x\n @x.deleter\n def x(self):\n  del self._x\nc = C()\nc._x = 1\ndel c.x\nprint(hasattr(c, '_x'))\n",
    "False"
);
crate::runtime_case!(
    property_doc,
    "class C:\n @property\n def x(self):\n  '''doc'''\n  return 1\nprint(C.x.__doc__)\n",
    "doc"
);
crate::runtime_case!(
    slots_restrict_dict,
    "class C:\n __slots__ = ('a',)\n def __init__(self):\n  self.a = 1\nprint(C().a)\n",
    "1"
);
crate::runtime_case!(
    slots_multiple,
    "class C:\n __slots__ = ('a', 'b')\nc = C()\nc.a = 1\nc.b = 2\nprint(c.b)\n",
    "2"
);
crate::runtime_case!(
    slots_inheritance,
    "class B:\n __slots__ = ('a',)\nclass D(B):\n __slots__ = ('b',)\nd = D()\nd.a = 1\nd.b = 2\nprint(d.a, d.b)\n",
    "1 2"
);
crate::runtime_case!(
    metaclass_type,
    "class M(type):\n pass\nclass C(metaclass=M):\n pass\nprint(type(C) is M)\n",
    "True"
);
crate::runtime_case!(
    metaclass_custom_name,
    "class M(type):\n def __new__(mcs, name, bases, ns):\n  return super().__new__(mcs, name, bases, ns)\nclass C(metaclass=M):\n pass\nprint(C.__name__)\n",
    "C"
);
crate::runtime_case!(
    init_subclass_hook,
    "class B:\n def __init_subclass__(cls, **kw):\n  cls.hooked = True\nclass D(B):\n pass\nprint(D.hooked)\n",
    "True"
);
crate::runtime_case!(
    class_getitem_generic,
    "class C:\n def __class_getitem__(cls, item):\n  return (cls, item)\nprint(C[int])\n",
    "(<class 'C'>, <class 'int'>)"
);
crate::runtime_case!(
    getattr_class,
    "class C:\n @classmethod\n def f(cls):\n  return 'cls'\nprint(C.f())\n",
    "cls"
);
crate::runtime_case!(
    descriptor_data_descriptor,
    "class D:\n def __get__(self, obj, owner):\n  return 1\n def __set__(self, obj, val):\n  pass\nclass C:\n x = D()\nprint(C.x)\n",
    "1"
);
crate::runtime_case!(
    descriptor_non_data,
    "class D:\n def __get__(self, obj, owner):\n  return 99\nclass C:\n x = D()\nc = C()\nc.__dict__['x'] = 1\nprint(c.x)\n",
    "1"
);
crate::runtime_case!(
    property_cached_pattern,
    "class C:\n def __init__(self):\n  self._cache = None\n @property\n def x(self):\n  if self._cache is None:\n   self._cache = 7\n  return self._cache\nprint(C().x)\n",
    "7"
);
crate::runtime_case!(
    staticmethod_no_descriptor_on_instance,
    "class C:\n @staticmethod\n def f():\n  return 3\nprint(C.f())\n",
    "3"
);
crate::runtime_case!(
    classmethod_descriptor,
    "class C:\n @classmethod\n def who(cls):\n  return cls.__name__\nprint(C.who())\n",
    "C"
);
crate::runtime_case!(
    getattr_magic,
    "class C:\n def __getattr__(self, name):\n  return 'missing'\nprint(C().xyz)\n",
    "missing"
);
crate::runtime_case!(
    setattr_magic,
    "class C:\n def __setattr__(self, name, val):\n  object.__setattr__(self, name, val * 2)\nc = C()\nc.x = 3\nprint(c.x)\n",
    "6"
);
crate::runtime_case!(
    delattr_magic,
    "class C:\n def __delattr__(self, name):\n  object.__delattr__(self, name)\nc = C()\nc.x = 1\ndel c.x\nprint(hasattr(c, 'x'))\n",
    "False"
);
crate::runtime_case!(
    dir_magic,
    "class C:\n def __dir__(self):\n  return ['custom']\nprint(C().__dir__())\n",
    "['custom']"
);
crate::runtime_case!(
    slots_no_dict,
    "class C:\n __slots__ = ()\nprint(hasattr(C(), '__dict__'))\n",
    "False"
);
crate::runtime_case!(
    metaclass_prepare,
    "class M(type):\n @classmethod\n def __prepare__(mcs, name, bases, **kw):\n  return {}\nclass C(metaclass=M):\n x = 1\nprint(C.x)\n",
    "1"
);
crate::runtime_case!(
    abstractmethod_property,
    "from abc import ABC, abstractmethod\nclass B(ABC):\n @property\n @abstractmethod\n def x(self):\n  pass\nprint(hasattr(B, 'x'))\n",
    "True"
);
crate::runtime_case!(
    descriptor_name_attr,
    "class D:\n def __get__(self, obj, owner):\n  return self\nclass C:\n x = D()\nprint(C.x is C.x)\n",
    "True"
);
crate::runtime_case!(
    property_member_descriptor,
    "class C:\n @property\n def x(self):\n  return 1\nprint(type(C.x).__name__)\n",
    "property"
);
crate::runtime_case!(
    classmethod_member,
    "class C:\n @classmethod\n def f(cls):\n  pass\nprint(type(C.f).__name__)\n",
    "method"
);
crate::runtime_case!(
    staticmethod_member,
    "class C:\n @staticmethod\n def f():\n  pass\nprint(type(C.f).__name__)\n",
    "function"
);
crate::runtime_case!(
    super_descriptor,
    "class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return super().f() + 1\nprint(D().f())\n",
    "2"
);
crate::runtime_case!(
    metaclass_instancecheck,
    "class M(type):\n def __instancecheck__(cls, instance):\n  return True\nclass C(metaclass=M):\n pass\nprint(isinstance(1, C))\n",
    "True"
);
crate::runtime_case!(
    metaclass_subclasscheck,
    "class M(type):\n def __subclasscheck__(cls, sub):\n  return True\nclass C(metaclass=M):\n pass\nprint(issubclass(int, C))\n",
    "True"
);
crate::runtime_case!(
    descriptor_owner_class,
    "class D:\n def __get__(self, obj, owner):\n  return owner.__name__\nclass C:\n x = D()\nprint(C().x)\n",
    "C"
);
crate::runtime_case!(
    slots_getattr_fallback,
    "class C:\n __slots__ = ('a',)\n def __init__(self):\n  self.a = 1\ntry:\n C().b\n print('ok')\nexcept AttributeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    property_setter_readonly,
    "class C:\n @property\n def x(self):\n  return 1\ntry:\n C().x = 2\n print('ok')\nexcept AttributeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    init_subclass_kwargs,
    "class B:\n def __init_subclass__(cls, flag=False, **kw):\n  cls.flag = flag\nclass D(B, flag=True):\n pass\nprint(D.flag)\n",
    "True"
);
crate::runtime_case!(
    descriptor_repr,
    "class D:\n def __get__(self, obj, owner):\n  return 1\nclass C:\n x = D()\nprint('D' in repr(C.__dict__['x']))\n",
    "True"
);
crate::runtime_case!(
    metaclass_call,
    "class M(type):\n def __call__(cls, *a, **k):\n  return 99\nclass C(metaclass=M):\n pass\nprint(C())\n",
    "99"
);
crate::runtime_case!(
    class_cell_closure,
    "def outer():\n x = 1\n class C:\n  def f(self):\n   return x\n return C\nprint(outer().f())\n",
    "1"
);
crate::runtime_case!(
    descriptor_set_name,
    "class D:\n def __set_name__(self, owner, name):\n  self.n = name\nclass C:\n x = D()\nprint(C().x.n)\n",
    "x"
);
crate::runtime_case!(
    enum_descriptor,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(E.A.name)\n",
    "A"
);

crate::compile_case!(
    metaclass_conflict,
    "class M1(type): pass\nclass M2(type): pass\ntry:\n class C(metaclass=M1, metaclass=M2): pass\nexcept TypeError: pass\n"
);
crate::compile_case!(
    slots_weakref,
    "class C:\n __slots__ = ('__weakref__', 'x')\n"
);
crate::compile_case!(
    descriptor_getset,
    "class D:\n def __get__(self, obj, owner): return 1\n def __set__(self, obj, val): pass\nclass C:\n x = D()\n"
);
crate::compile_case!(
    abc_register_virtual,
    "from abc import ABC\nclass B(ABC):\n pass\nclass C: pass\nB.register(C)\n"
);
crate::compile_case!(
    dataclass_with_slots,
    "from dataclasses import dataclass\n@dataclass(slots=True)\nclass P:\n x: int\n"
);
