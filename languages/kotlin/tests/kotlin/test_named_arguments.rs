use crate::helpers::run_prints;

#[test]
fn test_named_arguments_basic_ordering() {
    let out = run_prints(r#"
        fun format(label: String, count: Int): String {
            return label + ":" + count
        }
        fun main() {
            println(format(count = 2, label = "k"))
        }
    "#);
    assert_eq!(out, &["k:2"]);
}

#[test]
fn test_named_arguments_mixed_with_positional() {
    let out = run_prints(r#"
        fun make(a: Int, b: Int, c: Int): Int = a + b + c
        fun main() {
            println(make(1, c = 3, b = 2))
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_named_arguments_skips_middle_positionals() {
    let out = run_prints(r#"
        fun score(base: Int, bonus: Int, penalty: Int): Int = base + bonus - penalty
        fun main() {
            println(score(base = 10, penalty = 1, bonus = 2))
        }
    "#);
    assert_eq!(out, &["11"]);
}

#[test]
fn test_named_arguments_called_from_defaulted_fn() {
    let out = run_prints(r#"
        fun greet(prefix: String, name: String, suffix: String = "!"): String {
            return prefix + name + suffix
        }
        fun main() {
            println(greet(name = "k", prefix = "<", suffix = ">"))
            println(greet("[", "m"))
        }
    "#);
    assert_eq!(out, &["<k>", "[m!"]);
}

#[test]
fn test_named_arguments_all_defaults_can_be_overridden_by_name() {
    let out = run_prints(r#"
        fun build(prefix: String = "a", middle: String = "b", suffix: String = "c"): String {
            return prefix + middle + suffix
        }
        fun main() {
            println(build())
            println(build(suffix = "Z"))
            println(build(middle = "Y", prefix = "X"))
        }
    "#);
    assert_eq!(out, &["abc", "abZ", "XYc"]);
}

#[test]
fn test_named_arguments_single_named_on_vararg_function() {
    let out = run_prints(r#"
        fun join(prefix: String, vararg values: String, sep: String = ","): String {
            return prefix + values.joinToString(sep)
        }
        fun main() {
            println(join(prefix = "x", values = arrayOf("a", "b"), sep = ":"))
            println(join("x", "1", "2", sep = ";"))
        }
    "#);
    assert_eq!(out, &["x:a:b", "x1;2"]);
}

#[test]
fn test_named_arguments_in_method_calls() {
    let out = run_prints(r#"
        class Tagger {
            fun compose(head: String, value: String, tail: String): String = head + value + tail
        }
        fun main() {
            val t = Tagger()
            println(t.compose(head = "[", value = "v", tail = "]"))
        }
    "#);
    assert_eq!(out, &["[v]"]);
}

#[test]
fn test_named_arguments_object_factory_style() {
    let out = run_prints(r#"
        class Item(val a: Int, val b: Int)
        fun main() {
            val i = Item(a = 3, b = 4)
            println(i.a + i.b)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_named_arguments_in_secondary_constructor_call() {
    let out = run_prints(r#"
        class Box {
            val value: Int
            val tag: String
            constructor(value: Int, tag: String = "x") {
                this.value = value
                this.tag = tag
            }
        }
        fun main() {
            val x = Box(tag = "z", value = 9)
            println(x.value)
            println(x.tag)
        }
    "#);
    assert_eq!(out, &["9", "z"]);
}

#[test]
fn test_named_arguments_with_nullable_receiver() {
    let out = run_prints(r#"
        fun choose(first: String?, second: String = "b"): String {
            return (first ?: second)
        }
        fun main() {
            println(choose(first = null, second = "k"))
        }
    "#);
    assert_eq!(out, &["k"]);
}

#[test]
fn test_named_arguments_data_class_copy_named_args() {
    let out = run_prints(r#"
        data class User(val id: Int, val role: String)
        fun main() {
            val u = User(id = 2, role = "x")
            val v = u.copy(role = "admin")
            println(v.id)
            println(v.role)
        }
    "#);
    assert_eq!(out, &["2", "admin"]);
}

#[test]
fn test_named_arguments_lambda_argument_name_in_callable_reference() {
    let out = run_prints(r#"
        fun apply(label: String, fn: (String) -> String = { it }): String {
            return fn(label)
        }
        fun main() {
            println(apply(label = "x", fn = { "v:" + it }))
            println(apply(label = "y"))
        }
    "#);
    assert_eq!(out, &["v:x", "y"]);
}

#[test]
fn test_named_arguments_boolean_flags() {
    let out = run_prints(r#"
        fun pack(includeA: Boolean = true, includeB: Boolean = false): String {
            return (if (includeA) "A" else "") + (if (includeB) "B" else "")
        }
        fun main() {
            println(pack())
            println(pack(includeB = true))
            println(pack(includeA = false, includeB = true))
        }
    "#);
    assert_eq!(out, &["A", "AB", "B"]);
}

#[test]
fn test_named_arguments_uses_default_before_named_override() {
    let out = run_prints(r#"
        fun concat(a: String = "1", b: String, c: String = "3"): String {
            return a + b + c
        }
        fun main() {
            println(concat(b = "2"))
            println(concat(a = "A", b = "2", c = "C"))
        }
    "#);
    assert_eq!(out, &["123", "A2C"]);
}

#[test]
fn test_named_arguments_with_nested_function_call() {
    let out = run_prints(r#"
        fun outer(left: Int, right: Int): Int = left + right
        fun main() {
            fun inner(a: Int, b: Int, c: Int): Int = a + b + c
            println(outer(3, right = 4))
            println(inner(a = 1, b = 2, c = 3))
        }
    "#);
    assert_eq!(out, &["7", "6"]);
}

#[test]
fn test_named_arguments_in_extension_receiver_style() {
    let out = run_prints(r#"
        fun Int.scale(base: Int = 2, times: Int = 1): Int {
            return this * base + times
        }
        fun main() {
            println(3.scale(times = 5))
            println(3.scale(base = 4, times = 1))
        }
    "#);
    assert_eq!(out, &["11", "13"]);
}

#[test]
fn test_named_arguments_uses_named_when_invoking_overloaded() {
    let out = run_prints(r#"
        fun parse(value: String): String = "s" + value
        fun parse(value: Int): String = "i" + value
        fun main() {
            println(parse(value = "x"))
            println(parse(value = 7))
        }
    "#);
    assert_eq!(out, &["sx", "i7"]);
}

#[test]
fn test_named_arguments_named_and_default_in_chained_calls() {
    let out = run_prints(r#"
        fun base(x: Int = 1, y: Int = 2): Int = x + y
        fun scale(x: Int, y: Int = 2): Int = x * y
        fun main() {
            println(base(y = 9))
            println(scale(3, y = 5))
        }
    "#);
    assert_eq!(out, &["10", "15"]);
}

#[test]
fn test_named_arguments_with_type_inference_for_defaults() {
    let out = run_prints(r#"
        fun make(items: List<Int> = listOf(1, 2), label: String): String {
            return label + ":" + items.size
        }
        fun main() {
            println(make(label = "x"))
            println(make(items = listOf(1), label = "y"))
        }
    "#);
    assert_eq!(out, &["x:2", "y:1"]);
}

#[test]
fn test_named_arguments_boolean_with_shadowed_name() {
    let out = run_prints(r#"
        fun emit(enabled: Boolean = false, label: String = "off"): String {
            return if (enabled) "on:" + label else "off:" + label
        }
        fun main() {
            val enabled = true
            println(emit(enabled = enabled, label = "v"))
        }
    "#);
    assert_eq!(out, &["on:v"]);
}

#[test]
fn test_named_arguments_empty_string_defaults_named_call() {
    let out = run_prints(r#"
        fun wrap(prefix: String = "[", body: String, suffix: String = "]"): String {
            return prefix + body + suffix
        }
        fun main() {
            println(wrap(body = "x"))
            println(wrap(prefix = "<", body = "y", suffix = ">"))
        }
    "#);
    assert_eq!(out, &["[x]", "<y>"]);
}

#[test]
fn test_named_arguments_nested_defaults_and_name_scope() {
    let out = run_prints(r#"
        fun outer(tag: String, a: Int = 1, b: Int = 2): Int {
            return a + b
        }
        fun main() {
            println(outer("t", b = 6))
            println(outer(tag = "u", a = 2, b = 8))
        }
    "#);
    assert_eq!(out, &["7", "10"]);
}

#[test]
fn test_named_arguments_no_name_repetition() {
    let out = run_prints(r#"
        fun first(a: String, b: String): String = a + b
        fun main() {
            println(first(a = "a", b = "b"))
        }
    "#);
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_named_arguments_in_generic_functions() {
    let out = run_prints(r#"
        fun <T> pair(left: T, right: T): String = left.toString() + "," + right.toString()
        fun main() {
            println(pair<T=String>(left = "a", right = "b"))
            println(pair<Int>(left = 1, right = 2))
        }
    "#);
    assert_eq!(out, &["a,b", "1,2"]);
}

#[test]
fn test_named_arguments_chained_named_constructors() {
    let out = run_prints(r#"
        class Holder(val a: Int, val b: Int)
        class Container(val left: Holder, val title: String)
        fun main() {
            val c = Container(left = Holder(a = 1, b = 2), title = "t")
            println(c.left.a + c.left.b)
            println(c.title)
        }
    "#);
    assert_eq!(out, &["3", "t"]);
}

#[test]
fn test_named_arguments_method_with_receiver_and_defaults() {
    let out = run_prints(r#"
        fun String.pad(pre: String = "<", post: String = ">"): String {
            return pre + this + post
        }
        fun main() {
            println("x".pad(pre = "[", post = "]"))
            println("y".pad(post = "]"))
        }
    "#);
    assert_eq!(out, &["[x]", "<y>"]);
}

#[test]
fn test_named_arguments_named_arguments_internally_consistent() {
    let out = run_prints(r#"
        fun total(one: Int = 1, two: Int = 2, three: Int = 3): Int {
            return one + two + three
        }
        fun main() {
            println(total())
            println(total(two = 10))
            println(total(three = 7, one = 1, two = 2))
        }
    "#);
    assert_eq!(out, &["6", "12", "10"]);
}

#[test]
fn test_named_arguments_named_with_lambda_argument() {
    let out = run_prints(r#"
        fun map(value: Int, transform: (Int) -> Int = { it }, offset: Int = 0): Int {
            return transform(value) + offset
        }
        fun main() {
            println(map(value = 2, offset = 1))
            println(map(3, transform = { it * it }))
        }
    "#);
    assert_eq!(out, &["3", "9"]);
}

#[test]
fn test_named_arguments_nested_name_collision_guard() {
    let out = run_prints(r#"
        fun combine(a: String, b: String): String = a + b
        fun main() {
            val a = "x"
            println(combine(a = "u", b = a))
        }
    "#);
    assert_eq!(out, &["ux"]);
}

#[test]
fn test_named_arguments_named_call_for_unary_like() {
    let out = run_prints(r#"
        fun score(base: Int, delta: Int = 1): Int = base + delta
        fun main() {
            println(score(base = 9))
            println(score(base = 9, delta = 0))
        }
    "#);
    assert_eq!(out, &["10", "9"]);
}

#[test]
fn test_named_arguments_empty_call_path() {
    let out = run_prints(r#"
        fun pick(a: String, b: String = "x", c: String = "y"): String = a + b + c
        fun main() {
            println(pick(a = "1", b = "2", c = "3"))
        }
    "#);
    assert_eq!(out, &["123"]);
}
