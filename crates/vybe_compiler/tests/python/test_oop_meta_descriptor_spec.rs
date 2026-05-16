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

compile_case!(property_setter_compile, "class C:\n    def __init__(self):\n        self._x = 0\n    @property\n    def x(self):\n        return self._x\n    @x.setter\n    def x(self, value):\n        self._x = value\n");
compile_case!(property_deleter_compile, "class C:\n    def __init__(self):\n        self._x = 0\n    @property\n    def x(self):\n        return self._x\n    @x.deleter\n    def x(self):\n        del self._x\n");
compile_case!(descriptor_set_name_compile, "class D:\n    def __set_name__(self, owner, name):\n        self.name = name\nclass C:\n    value = D()\n");
compile_case!(init_subclass_compile, "class Base:\n    def __init_subclass__(cls, flag=False, **kwargs):\n        super().__init_subclass__(**kwargs)\nclass Child(Base, flag=True):\n    pass\n");
compile_case!(class_getitem_compile, "class Box:\n    def __class_getitem__(cls, item):\n        return cls\nT = Box[int]\n");
compile_case!(new_method_compile, "class C:\n    def __new__(cls, *args, **kwargs):\n        return super().__new__(cls)\n");
compile_case!(prepare_metaclass_compile, "class Meta(type):\n    @classmethod\n    def __prepare__(mcls, name, bases):\n        return {}\nclass C(metaclass=Meta):\n    pass\n");
compile_case!(metaclass_new_compile, "class Meta(type):\n    def __new__(mcls, name, bases, ns):\n        return super().__new__(mcls, name, bases, ns)\n");
compile_case!(metaclass_call_compile, "class Meta(type):\n    def __call__(cls, *args, **kwargs):\n        return super().__call__(*args, **kwargs)\n");
compile_case!(mro_access_compile, "class A: pass\nclass B(A): pass\norder = B.__mro__\n");
compile_case!(mro_method_compile, "class A: pass\nclass B(A): pass\norder = B.mro()\n");
compile_case!(super_explicit_compile, "class A:\n    def f(self):\n        return 1\nclass B(A):\n    def f(self):\n        return super(B, self).f()\n");
compile_case!(super_in_classmethod_compile, "class A:\n    @classmethod\n    def f(cls):\n        return 1\nclass B(A):\n    @classmethod\n    def f(cls):\n        return super().f()\n");
compile_case!(super_in_staticmethod_compile, "class A:\n    @staticmethod\n    def f():\n        return 1\nclass B(A):\n    @staticmethod\n    def g():\n        return A.f()\n");
compile_case!(slots_with_property_compile, "class C:\n    __slots__ = ('_x',)\n    @property\n    def x(self):\n        return self._x\n");
compile_case!(annotations_dict_compile, "class C:\n    x: int\n    y: str\nann = C.__annotations__\n");
compile_case!(class_dict_compile, "class C:\n    x = 1\nd = C.__dict__\n");
compile_case!(instance_dict_compile, "class C:\n    pass\nc = C()\nd = c.__dict__\n");
compile_case!(abstractmethod_compile, "from abc import ABC, abstractmethod\nclass Base(ABC):\n    @abstractmethod\n    def run(self):\n        pass\n");
compile_case!(abstractproperty_compile, "from abc import ABC, abstractmethod\nclass Base(ABC):\n    @property\n    @abstractmethod\n    def value(self):\n        pass\n");
compile_case!(abc_register_compile, "from abc import ABC\nclass Base(ABC):\n    pass\nBase.register(tuple)\n");
compile_case!(subclasshook_compile, "from abc import ABCMeta\nclass Base(metaclass=ABCMeta):\n    @classmethod\n    def __subclasshook__(cls, C):\n        return NotImplemented\n");
compile_case!(datamodel_repr_compile, "class C:\n    def __repr__(self):\n        return 'C()'\n    def __str__(self):\n        return 'c'\n");
runtime_case!(property_runtime_basic, "class C:\n    def __init__(self):\n        self._x = 5\n    @property\n    def x(self):\n        return self._x\nprint(C().x)\n", "5");
runtime_case!(classmethod_runtime_basic, "class C:\n    @classmethod\n    def name(cls):\n        return 'C'\nprint(C.name())\n", "C");
runtime_case!(staticmethod_runtime_basic, "class C:\n    @staticmethod\n    def add(a, b):\n        return a + b\nprint(C.add(2, 3))\n", "5");
runtime_case!(super_runtime_chain, "class A:\n    def f(self):\n        return 1\nclass B(A):\n    def f(self):\n        return super().f() + 1\nprint(B().f())\n", "2");
runtime_case!(instance_dict_runtime, "class C:\n    pass\nc = C()\nc.x = 9\nprint(c.__dict__['x'])\n", "9");
runtime_case!(class_attr_runtime, "class C:\n    y = 7\nprint(C.__dict__['y'])\n", "7");
compile_case!(multiple_inheritance_super_compile, "class A: pass\nclass B(A): pass\nclass C(A): pass\nclass D(B, C):\n    def f(self):\n        return super().f()\n");