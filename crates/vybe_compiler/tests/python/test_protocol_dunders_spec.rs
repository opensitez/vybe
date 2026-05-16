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

compile_case!(dunder_sub_compile, "class Vec:\n    def __sub__(self, other):\n        return self\n");
compile_case!(dunder_mul_compile, "class Vec:\n    def __mul__(self, other):\n        return self\n");
compile_case!(dunder_truediv_compile, "class Vec:\n    def __truediv__(self, other):\n        return self\n");
compile_case!(dunder_floordiv_compile, "class Vec:\n    def __floordiv__(self, other):\n        return self\n");
compile_case!(dunder_mod_compile, "class Vec:\n    def __mod__(self, other):\n        return self\n");
compile_case!(dunder_pow_compile, "class Vec:\n    def __pow__(self, other):\n        return self\n");
compile_case!(dunder_lt_compile, "class Box:\n    def __lt__(self, other):\n        return True\n");
compile_case!(dunder_le_compile, "class Box:\n    def __le__(self, other):\n        return True\n");
compile_case!(dunder_gt_compile, "class Box:\n    def __gt__(self, other):\n        return False\n");
compile_case!(dunder_ge_compile, "class Box:\n    def __ge__(self, other):\n        return False\n");
compile_case!(dunder_neg_compile, "class Num:\n    def __neg__(self):\n        return self\n");
compile_case!(dunder_pos_compile, "class Num:\n    def __pos__(self):\n        return self\n");
compile_case!(dunder_abs_compile, "class Num:\n    def __abs__(self):\n        return 0\n");
compile_case!(dunder_invert_compile, "class Mask:\n    def __invert__(self):\n        return self\n");
runtime_case!(dunder_call_runtime, "class Adder:\n    def __call__(self, x, y):\n        return x + y\nadd = Adder()\nprint(add(2, 3))\n", "5");
runtime_case!(dunder_bool_runtime, "class Flag:\n    def __bool__(self):\n        return True\nprint(bool(Flag()))\n", "true");
compile_case!(dunder_getattr_compile, "class Proxy:\n    def __getattr__(self, name):\n        return 0\n");
compile_case!(dunder_setattr_compile, "class Proxy:\n    def __setattr__(self, name, value):\n        super().__setattr__(name, value)\n");
compile_case!(dunder_delattr_compile, "class Proxy:\n    def __delattr__(self, name):\n        super().__delattr__(name)\n");
compile_case!(dunder_getattribute_compile, "class Proxy:\n    def __getattribute__(self, name):\n        return super().__getattribute__(name)\n");
compile_case!(slots_compile, "class Point:\n    __slots__ = ('x', 'y')\n");
compile_case!(descriptor_get_compile, "class D:\n    def __get__(self, obj, owner):\n        return 1\n");
compile_case!(descriptor_set_compile, "class D:\n    def __set__(self, obj, value):\n        pass\n");
compile_case!(descriptor_delete_compile, "class D:\n    def __delete__(self, obj):\n        pass\n");
compile_case!(dunder_iadd_compile, "class Counter:\n    def __iadd__(self, other):\n        return self\n");
compile_case!(dunder_isub_compile, "class Counter:\n    def __isub__(self, other):\n        return self\n");
compile_case!(dunder_imul_compile, "class Counter:\n    def __imul__(self, other):\n        return self\n");
compile_case!(dunder_iter_compile, "class Seq:\n    def __iter__(self):\n        return self\n");
compile_case!(dunder_next_compile, "class Seq:\n    def __next__(self):\n        raise StopIteration\n");
compile_case!(dunder_contains_compile, "class Bag:\n    def __contains__(self, item):\n        return False\n");