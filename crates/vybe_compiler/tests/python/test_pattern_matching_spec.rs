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

compile_case!(match_sequence_exact_compile, "match [1, 2]:\n    case [1, 2]:\n        pass\n");
compile_case!(match_sequence_star_compile, "match [1, 2, 3, 4]:\n    case [1, *rest]:\n        pass\n");
compile_case!(match_tuple_compile, "match (1, 2):\n    case (1, x):\n        print(x)\n");
compile_case!(match_nested_sequence_compile, "match [1, [2, 3]]:\n    case [1, [a, b]]:\n        print(a, b)\n");
compile_case!(match_mapping_exact_compile, "match {'x': 1, 'y': 2}:\n    case {'x': 1, 'y': y}:\n        print(y)\n");
compile_case!(match_mapping_rest_compile, "match {'x': 1, 'y': 2}:\n    case {'x': x, **rest}:\n        print(x, rest)\n");
compile_case!(match_class_positional_compile, "class Point:\n    __match_args__ = ('x', 'y')\nmatch Point():\n    case Point(1, y):\n        print(y)\n");
compile_case!(match_class_keyword_compile, "class Point:\n    pass\nmatch Point():\n    case Point(x=1, y=y):\n        print(y)\n");
compile_case!(match_as_pattern_compile, "match [1, 2]:\n    case [1, x] as whole:\n        print(x, whole)\n");
runtime_case!(match_capture_pattern_runtime, "value = [1, 2]\nmatch value:\n    case [1, x]:\n        print(x)\n", "2");
compile_case!(match_true_compile, "match True:\n    case True:\n        pass\n");
compile_case!(match_false_compile, "match False:\n    case False:\n        pass\n");
compile_case!(match_negative_number_compile, "match -5:\n    case -5:\n        pass\n");
compile_case!(match_float_compile, "match 3.14:\n    case 3.14:\n        pass\n");
compile_case!(match_bytes_compile, "match b'hi':\n    case b'hi':\n        pass\n");
compile_case!(match_sequence_or_compile, "match [1, 2]:\n    case [1, 2] | [2, 1]:\n        pass\n");
compile_case!(match_mapping_guard_compile, "match {'x': 10}:\n    case {'x': x} if x > 5:\n        pass\n");
compile_case!(match_class_guard_compile, "class Box:\n    pass\nmatch Box():\n    case Box() as box if box is not None:\n        pass\n");
compile_case!(match_nested_class_compile, "class Node:\n    pass\nmatch Node():\n    case Node(left=Node(), right=_):\n        pass\n");
compile_case!(match_star_middle_compile, "match [1, 2, 3, 4]:\n    case [first, *middle, last]:\n        print(first, middle, last)\n");
compile_case!(match_star_end_compile, "match [1, 2, 3]:\n    case [first, *rest]:\n        print(first, rest)\n");
compile_case!(match_tuple_or_compile, "match (1, 2):\n    case (1, 2) | (2, 1):\n        pass\n");
compile_case!(match_dict_literal_key_compile, "match {'kind': 'ok', 'value': 1}:\n    case {'kind': 'ok', 'value': value}:\n        print(value)\n");
compile_case!(match_subject_call_compile, "def make():\n    return [1, 2]\nmatch make():\n    case [1, x]:\n        print(x)\n");
compile_case!(match_wildcard_after_sequence_compile, "match [1, 2, 3]:\n    case [1, *_]:\n        pass\n");
compile_case!(match_as_with_or_compile, "match 1:\n    case (1 | 2) as value:\n        print(value)\n");
compile_case!(match_list_of_tuples_compile, "match [('a', 1), ('b', 2)]:\n    case [('a', x), *rest]:\n        print(x, rest)\n");
compile_case!(match_mapping_double_star_compile, "match {'a': 1, 'b': 2}:\n    case {'a': a, **rest}:\n        print(a, rest)\n");
compile_case!(match_dotted_value_compile, "class Colors:\n    RED = 'red'\ncolor = 'red'\nmatch color:\n    case Colors.RED:\n        pass\n");
runtime_case!(match_sequence_runtime, "value = [1, 2, 3]\nmatch value:\n    case [1, 2, *rest]:\n        print(rest[0])\n", "3");