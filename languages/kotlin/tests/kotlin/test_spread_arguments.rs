use crate::helpers::run_prints;

#[test]
fn test_spread_with_ints() {
    let out = run_prints(r#"
        fun join(prefix: String, vararg values: Int): String {
            return prefix + values.joinToString(",")
        }
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            println(join("v", *nums))
        }
    "#);
    assert_eq!(out, &["v1,2,3"]);
}

#[test]
fn test_spread_with_no_varargs() {
    let out = run_prints(r#"
        fun join(prefix: String, vararg values: Int): String {
            return if (values.isEmpty()) "empty" else prefix + values.joinToString(";")
        }
        fun main() {
            println(join("x"))
        }
    "#);
    assert_eq!(out, &["empty"]);
}

#[test]
fn test_spread_array_plus_vararg() {
    let out = run_prints(r#"
        fun combine(base: String, vararg tags: String): String = base + tags.joinToString("|")
        fun main() {
            val tags = arrayOf("a", "b")
            println(combine("x", "c", *tags, "d"))
        }
    "#);
    assert_eq!(out, &["xa|b|d|c"]);
}

#[test]
fn test_spread_with_empty_array() {
    let out = run_prints(r#"
        fun join(vararg values: Int): Int = values.size
        fun main() {
            val empty = intArrayOf()
            println(join(*empty))
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_spread_in_nested_calls() {
    let out = run_prints(r#"
        fun add(a: Int, b: Int, c: Int): Int = a + b + c
        fun main() {
            val head = intArrayOf(1, 2)
            println(add(3, *head))
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_spread_with_mixed_types_disallowed() {
    let out = run_prints(r#"
        fun collect(vararg values: String): String = values.joinToString(":")
        fun main() {
            val a = arrayOf("x", "y")
            println(collect(*a))
        }
    "#);
    assert_eq!(out, &["x:y"]);
}

#[test]
fn test_spread_with_boxed_arrays() {
    let out = run_prints(r#"
        fun sumAll(values: IntArray): Int {
            var total = 0
            for (v in values) total += v
            return total
        }
        fun main() {
            val a = intArrayOf(1, 2, 3)
            println(sumAll(a))
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_spread_with_object_array() {
    let out = run_prints(r#"
        fun describe(prefix: String, vararg values: Any): String {
            return prefix + values.size
        }
        fun main() {
            val vals: Array<Any> = arrayOf(1, "x", true)
            println(describe("n", *vals))
        }
    "#);
    assert_eq!(out, &["n3"]);
}

#[test]
fn test_spread_with_double_array() {
    let out = run_prints(r#"
        fun sum(a: Double, b: Double, c: Double): Double = a + b + c
        fun main() {
            val arr = doubleArrayOf(1.0, 2.0)
            println(sum(3.0, *arr))
        }
    "#);
    assert_eq!(out, &["6.0"]);
}

#[test]
fn test_spread_string_builder_join() {
    let out = run_prints(r#"
        fun main() {
            val parts = arrayOf("a", "b")
            val all = arrayOf("s", *parts, "t")
            println(all.joinToString("-"))
        }
    "#);
    assert_eq!(out, &["s-a-b-t"]);
}

#[test]
fn test_spread_char_sequence_vararg() {
    let out = run_prints(r#"
        fun concat(vararg items: Char): String {
            return items.joinToString("")
        }
        fun main() {
            val head = charArrayOf('x', 'y')
            println(concat(*head, 'z'))
        }
    "#);
    assert_eq!(out, &["xyz"]);
}

#[test]
fn test_spread_with_defaulted_vararg() {
    let out = run_prints(r#"
        fun prefix(base: String, vararg values: Int = intArrayOf(9)): String {
            return base + values.joinToString(".")
        }
        fun main() {
            println(prefix("a"))
            println(prefix("a", 1, 2))
        }
    "#);
    assert_eq!(out, &["a9", "a1.2"]);
}

#[test]
fn test_spread_with_list_to_array_conversion() {
    let out = run_prints(r#"
        fun sum(base: Int, values: IntArray): Int {
            var total = base
            for (v in values) total += v
            return total
        }
        fun main() {
            val items = listOf(1, 2, 3).toIntArray()
            println(sum(1, items))
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_spread_long_array() {
    let out = run_prints(r#"
        fun joinAll(prefix: String, vararg values: Long): String {
            return prefix + values.joinToString(",")
        }
        fun main() {
            val nums = longArrayOf(5L, 6L)
            println(joinAll("L", *nums))
        }
    "#);
    assert_eq!(out, &["L5,6"]);
}

#[test]
fn test_spread_calling_with_array_plus_literals() {
    let out = run_prints(r#"
        fun show(vararg values: String): String = values.joinToString(",")
        fun main() {
            val a = arrayOf("b", "c")
            println(show("a", *a, "d"))
        }
    "#);
    assert_eq!(out, &["a,b,c,d"]);
}

#[test]
fn test_spread_immutable_array_copy() {
    let out = run_prints(r#"
        fun copy(base: Int, values: IntArray): Int {
            return base + values.sum()
        }
        fun main() {
            val values = intArrayOf(2, 4)
            println(copy(3, values))
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_spread_boolean_vararg() {
    let out = run_prints(r#"
        fun allTrue(vararg values: Boolean): Boolean {
            for (v in values) if (!v) return false
            return true
        }
        fun main() {
            val flags = booleanArrayOf(true, true, false)
            println(allTrue(*flags))
            val flags2 = booleanArrayOf(true, true)
            println(allTrue(*flags2))
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_spread_nested_vararg_calls() {
    let out = run_prints(r#"
        fun outer(prefix: String, vararg values: String): String = prefix + values.joinToString(":")
        fun build(base: Array<String>): String {
            return outer("p", *base)
        }
        fun main() {
            println(build(arrayOf("x", "y")))
        }
    "#);
    assert_eq!(out, &["px:y"]);
}

#[test]
fn test_spread_with_vararg_reference() {
    let out = run_prints(r#"
        fun total(vararg values: Int): Int {
            return values.size
        }
        fun main() {
            val fnRef = ::total
            val nums = intArrayOf(1,2,3)
            println(fnRef(*nums))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_spread_in_class_method() {
    let out = run_prints(r#"
        class Acc {
            fun join(vararg values: Int): Int = values.size
        }
        fun main() {
            val a = Acc()
            println(a.join(1,2))
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_spread_with_zero_length_reference() {
    let out = run_prints(r#"
        fun join(vararg values: String): String = values.joinToString(",")
        fun main() {
            val empty = arrayOf<String>()
            println(join(*empty))
        }
    "#);
    assert_eq!(out, &[""]);
}

#[test]
fn test_spread_multiple_arrays_to_vararg() {
    let out = run_prints(r#"
        fun join(values: IntArray): Int {
            var total = 0
            for (v in values) total += v
            return total
        }
        fun sum(prefix: String, vararg values: Int): Int {
            return prefix.length + values.sum()
        }
        fun main() {
            val a = intArrayOf(1, 2)
            val b = intArrayOf(3, 4)
            println(sum("x", *a, *b))
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_spread_vararg_with_trailing_literals() {
    let out = run_prints(r#"
        fun join(prefix: String, vararg values: String): String = prefix + values.joinToString(".")
        fun main() {
            val mid = arrayOf("b", "c")
            println(join("a", *mid, "d"))
        }
    "#);
    assert_eq!(out, &["a.b.c.d"]);
}

#[test]
fn test_spread_between_calls() {
    let out = run_prints(r#"
        fun wrap(base: String, values: IntArray): String = base + values.joinToString(",")
        fun sink(prefix: String, vararg values: Int): Int {
            return values.sum() + prefix.length
        }
        fun main() {
            val values = intArrayOf(2, 3, 4)
            println(wrap("p", values))
            println(sink("x", *values))
        }
    "#);
    assert_eq!(out, &["p2,3,4", "8"]);
}

#[test]
fn test_spread_in_extension_function_call() {
    let out = run_prints(r#"
        fun String.tagged(vararg values: String): String = this + values.joinToString(":")
        fun main() {
            val tags = arrayOf("a", "b", "c")
            println("base:".tagged(*tags))
        }
    "#);
    assert_eq!(out, &["base:a:b:c"]);
}

#[test]
fn test_spread_through_vararg_array_param() {
    let out = run_prints(r#"
        fun flatten(prefix: String, values: IntArray): String = prefix + values.joinToString("|")
        fun main() {
            val a = intArrayOf(1)
            val b = intArrayOf(2, 3)
            println(flatten("x", intArrayOf(*a, *b)))
        }
    "#);
    assert_eq!(out, &["x1|2|3"]);
}

#[test]
fn test_spread_in_lambda_capture() {
    let out = run_prints(r#"
        fun total(vararg values: Int): Int = values.sum()
        fun main() {
            val source = intArrayOf(5, 6)
            val runTotal = { arr: IntArray -> total(*arr) }
            println(runTotal(source))
        }
    "#);
    assert_eq!(out, &["11"]);
}

#[test]
fn test_spread_char_array_to_string_vararg() {
    let out = run_prints(r#"
        fun pack(prefix: String, vararg values: Char): String = prefix + values.joinToString("")
        fun main() {
            val chars = charArrayOf('h', 'i')
            println(pack("say:", *chars))
        }
    "#);
    assert_eq!(out, &["say:hi"]);
}

#[test]
fn test_spread_vararg_reference_within_class() {
    let out = run_prints(r#"
        class Combiner {
            fun build(prefix: String, vararg values: Int): String {
                return prefix + values.joinToString(",")
            }
        }
        fun main() {
            val c = Combiner()
            val a = intArrayOf(7, 8)
            println(c.build("n", *a))
        }
    "#);
    assert_eq!(out, &["n7,8"]);
}

#[test]
fn test_spread_with_default_initializer_array() {
    let out = run_prints(r#"
        fun values(base: String = "v", vararg entries: Int = intArrayOf(1)): String {
            return base + entries.joinToString("|")
        }
        fun main() {
            val preset = intArrayOf()
            println(values())
            println(values("x", *preset))
        }
    "#);
    assert_eq!(out, &["v1", "x"]);
}
