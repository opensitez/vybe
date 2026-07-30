use crate::helpers::run_prints;

#[test]
fn test_typealias_for_aliasing_simple_type() {
    let out = run_prints(r#"
        typealias Text = String

        fun main() {
            val value: Text = "ok"
            println(value)
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_typealias_for_class_constructor_arguments() {
    let out = run_prints(r#"
        data class PairValue(val id: Int, val label: String)
        typealias Entry = PairValue

        fun main() {
            val item: Entry = Entry(3, "x")
            println(item.id)
            println(item.label)
        }
    "#);
    assert_eq!(out, &["3", "x"]);
}

#[test]
fn test_typealias_for_function_type() {
    let out = run_prints(r#"
        typealias Reducer = (Int, Int) -> Int

        fun combine(value: Int, other: Int, op: Reducer): Int {
            return op(value, other)
        }

        fun main() {
            val result = combine(4, 5, { a, b -> a + b })
            println(result)
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_typealias_for_generic_type() {
    let out = run_prints(r#"
        typealias BoxOfInt = MutableList<Int>

        fun main() {
            val values: BoxOfInt = mutableListOf(1, 2, 3)
            values.add(4)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_typealias_for_union_style_nesting() {
    let out = run_prints(r#"
        typealias Name = String
        typealias Named = Name

        fun main() {
            val value: Named = "x"
            println(value)
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_typealias_generic_holder_round_trip() {
    let out = run_prints(r#"
        class Holder<T>(val value: T)
        typealias StringHolder = Holder<String>

        fun main() {
            val item: StringHolder = StringHolder("done")
            println(item.value)
        }
    "#);
    assert_eq!(out, &["done"]);
}

#[test]
fn test_typealias_does_not_duplicate_runtime_identity() {
    let out = run_prints(r#"
        typealias Label = String

        class Box(var value: Label)

        fun main() {
            val first = Box("a")
            val second: Label = "a"
            first.value = second
            println(first.value)
        }
    "#);
    assert_eq!(out, &["a"]);
}

#[test]
fn test_typealias_for_map_type() {
    let out = run_prints(r#"
        typealias StringNumberMap = Map<String, Int>

        fun main() {
            val values: StringNumberMap = mapOf("a" to 1, "b" to 2)
            println(values["a"])
            println(values["c"] == null)
        }
    "#);
    assert_eq!(out, &["1", "true"]);
}

#[test]
fn test_local_typealias_scopes_to_block() {
    let out = run_prints(r#"
        fun make(): String {
            typealias LocalText = String
            val value: LocalText = "block"
            return value
        }

        fun main() {
            println(make())
        }
    "#);
    assert_eq!(out, &["block"]);
}

#[test]
fn test_typealias_with_array_type() {
    let out = run_prints(r#"
        typealias IntArrayLike = Array<Int>

        fun main() {
            val values: IntArrayLike = arrayOf(1, 2)
            println(values.size)
            println(values[0] + values[1])
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_typealias_for_tuple_like_type() {
    let out = run_prints(r#"
        typealias PairText = Pair<Int, String>

        fun main() {
            val value: PairText = Pair(4, "x")
            println(value.first)
            println(value.second)
        }
    "#);
    assert_eq!(out, &["4", "x"]);
}
