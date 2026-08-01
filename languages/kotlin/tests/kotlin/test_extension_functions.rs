use crate::helpers::run_prints;

#[test]
fn test_extension_function_on_primitive() {
    let out = run_prints(
        r#"
        fun Int.incremented(): Int = this + 1

        fun main() {
            println(3.incremented())
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_extension_function_with_multiple_parameters() {
    let out = run_prints(
        r#"
        fun Int.add(value: Int, scale: Int): Int = (this + value) * scale

        fun main() {
            println(2.add(3, 4))
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_extension_function_on_class_instance() {
    let out = run_prints(
        r#"
        class Box(val value: Int)

        fun Box.labeled(prefix: String): String = prefix + ":" + value

        fun main() {
            println(Box(7).labeled("v"))
        }
    "#,
    );
    assert_eq!(out, &["v:7"]);
}

#[test]
fn test_extension_property_getter() {
    let out = run_prints(
        r#"
        class Point(val x: Int, val y: Int)

        val Point.sum: Int
            get() = x + y

        fun main() {
            println(Point(2, 5).sum)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_extension_property_with_setter_like_behavior() {
    let out = run_prints(
        r#"
        class Holder(var value: Int)

        var Holder.doubled: Int
            get() = value * 2
            set(next) { value = next / 2 }

        fun main() {
            val holder = Holder(3)
            holder.doubled = 10
            println(holder.value)
            println(holder.doubled)
        }
    "#,
    );
    assert_eq!(out, &["5", "10"]);
}

#[test]
fn test_extension_function_for_nullable_receiver() {
    let out = run_prints(
        r#"
        fun Int?.orZero(): Int = this ?: 0

        fun main() {
            val value: Int? = null
            val second: Int? = 7
            println(value.orZero())
            println(second.orZero())
        }
    "#,
    );
    assert_eq!(out, &["0", "7"]);
}

#[test]
fn test_local_extension_function_scope() {
    let out = run_prints(
        r#"
        fun main() {
            fun String.shout(): String = this.uppercase()
            fun use(value: String): String = value.shout()
            println(use("go"))
        }
    "#,
    );
    assert_eq!(out, &["GO"]);
}

#[test]
fn test_overload_resolution_between_extension_and_member() {
    let out = run_prints(
        r#"
        class Box {
            fun value(): Int = 1
        }

        fun Box.value(): Int = 4

        fun main() {
            println(Box().value())
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_generic_extension_transform() {
    let out = run_prints(
        r#"
        fun <T> List<T>.wrapCount(): String = "count=" + this.size

        fun main() {
            println(listOf(1, 2, 3).wrapCount())
            println(listOf("a").wrapCount())
        }
    "#,
    );
    assert_eq!(out, &["count=3", "count=1"]);
}

#[test]
fn test_extension_on_generic_with_bounds() {
    let out = run_prints(
        r#"
        fun <T : Number> T.asIntText(): Int = this.toInt()

        fun main() {
            println(4.9.asIntText())
            println(7.asIntText())
        }
    "#,
    );
    assert_eq!(out, &["4", "7"]);
}

#[test]
fn test_extension_function_chain_with_let() {
    let out = run_prints(
        r#"
        fun String.repeatPrefix(prefix: String, count: Int): String = prefix.repeat(count) + this

        fun main() {
            val value = "k"
                .repeatPrefix("a", 3)
                .repeatPrefix("b", 2)
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["bbaaaak"]);
}

#[test]
fn test_extension_function_with_default_parameter() {
    let out = run_prints(
        r#"
        fun String.wrap(prefix: String = "x", suffix: String = "!"): String {
            return prefix + this + suffix
        }

        fun main() {
            println("a".wrap())
            println("a".wrap("z"))
            println("a".wrap("z", "?"))
        }
    "#,
    );
    assert_eq!(out, &["xa!", "za!", "za?"]);
}

#[test]
fn test_extension_function_on_any_type() {
    let out = run_prints(
        r#"
        fun Any.described(): String = when (this) {
            is Int -> "Int"
            is String -> "String"
            else -> "Any"
        }

        class Item

        fun main() {
            println(3.described())
            println("x".described())
            println(Item().described())
        }
    "#,
    );
    assert_eq!(out, &["Int", "String", "Any"]);
}

#[test]
fn test_extension_property_with_nullable_receiver() {
    let out = run_prints(
        r#"
        val Int?.orZero: Int
            get() = this ?: 0

        fun main() {
            val left: Int? = null
            val right: Int? = 12
            println(left.orZero)
            println(right.orZero)
        }
    "#,
    );
    assert_eq!(out, &["0", "12"]);
}

#[test]
fn test_extension_receiver_shadowing_from_scope() {
    let out = run_prints(
        r#"
        fun String.show(): String = "global-" + this

        fun main() {
            fun String.show(): String = "local-" + this
            println("x".show())
        }
    "#,
    );
    assert_eq!(out, &["local-x"]);
}

#[test]
fn test_extension_function_with_vararg_receiver() {
    let out = run_prints(
        r#"
        fun Int.joinWith(vararg values: Int): String {
            var total = this
            for (value in values) {
                total += value
            }
            return total.toString()
        }

        fun main() {
            println(1.joinWith(2, 3, 4))
            println(0.joinWith())
        }
    "#,
    );
    assert_eq!(out, &["10", "0"]);
}

#[test]
fn test_extension_to_collection_can_transform() {
    let out = run_prints(
        r#"
        fun <T> Collection<T>.asTagged(tag: String): String {
            return tag + ":" + this.size
        }

        fun main() {
            println(listOf(1, 2, 3).asTagged("count"))
            println(setOf("a").asTagged("single"))
        }
    "#,
    );
    assert_eq!(out, &["count:3", "single:1"]);
}

#[test]
fn test_extension_function_for_sequence_order() {
    let out = run_prints(
        r#"
        fun Int.next(): Int = this + 1
        fun Int.prev(): Int = this - 1

        fun main() {
            val base = 4
            println(base.next().prev().next())
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_extension_function_can_return_receiver_again() {
    let out = run_prints(
        r#"
        fun String.trimAndRepeat(times: Int): String {
            val clean = this.trim()
            return clean.repeat(times)
        }

        fun main() {
            println("  x ".trimAndRepeat(3))
            println("z".trimAndRepeat(1))
        }
    "#,
    );
    assert_eq!(out, &["xxx", "z"]);
}

#[test]
fn test_extension_in_generic_bounded_receiver() {
    let out = run_prints(
        r#"
        fun <T> T.describeIfString(default: String): String where T : Any? {
            return this?.toString() ?: default
        }

        fun main() {
            val value: String? = null
            println("hello".describeIfString("none"))
            println(value.describeIfString("none"))
        }
    "#,
    );
    assert_eq!(out, &["hello", "none"]);
}

#[test]
fn test_extension_property_can_be_computed_multiple_times() {
    let out = run_prints(
        r#"
        var Int.squareCount: Int
            get() = this * this
            set(value) {}

        fun main() {
            val value = 4
            println(value.squareCount)
            println(value.squareCount)
        }
    "#,
    );
    assert_eq!(out, &["16", "16"]);
}

#[test]
fn test_infix_extension_function_supports_chained_usage() {
    let out = run_prints(
        r#"
        infix fun String.tagWith(prefix: String): String = prefix + this

        fun main() {
            val result = "world" tagWith "hello-" tagWith "!"
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["!hello-world"]);
}

#[test]
fn test_extension_function_on_number_list_maps_all_values() {
    let out = run_prints(
        r#"
        fun List<Int>.doubleAndSum(): Int {
            var total = 0
            for (value in this) {
                total += value * 2
            }
            return total
        }

        fun main() {
            println(listOf(1, 2, 3).doubleAndSum())
            println(listOf(10).doubleAndSum())
        }
    "#,
    );
    assert_eq!(out, &["12", "20"]);
}

#[test]
fn test_extension_function_on_int_array_returns_product() {
    let out = run_prints(
        r#"
        fun IntArray.product(): Int {
            var total = 1
            for (value in this) {
                total *= value
            }
            return total
        }

        fun main() {
            println(intArrayOf(2, 3, 4).product())
        }
    "#,
    );
    assert_eq!(out, &["24"]);
}

#[test]
fn test_extension_property_on_map_reports_key_count() {
    let out = run_prints(
        r#"
        val Map<String, Int>.keyText: String
            get() = keys.joinToString("|")

        fun main() {
            val values = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            println(values.keyText)
            println(values.keyText)
        }
    "#,
    );
    assert_eq!(out, &["a|b|c", "a|b|c"]);
}

#[test]
fn test_extension_function_for_boolean_returns_numeric_projection() {
    let out = run_prints(
        r#"
        fun Boolean.intValue(): Int = if (this) 1 else 0

        fun main() {
            println(true.intValue())
            println(false.intValue())
        }
    "#,
    );
    assert_eq!(out, &["1", "0"]);
}

#[test]
fn test_extension_function_with_receiver_parameter() {
    let out = run_prints(
        r#"
        fun Int.addWithOffset(offset: Int, label: String): String {
            return label + (this + offset).toString()
        }

        fun main() {
            println(2.addWithOffset(3, "x"))
        }
    "#,
    );
    assert_eq!(out, &["x5"]);
}

#[test]
fn test_extension_function_on_pair_exposes_projection_like_accessor() {
    let out = run_prints(
        r#"
        fun Pair<Int, Int>.delta(): Int = second - first

        fun main() {
            val value = Pair(4, 9)
            println(value.delta())
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_extension_function_can_be_inference_targeted_by_generic_constraint() {
    let out = run_prints(
        r#"
        fun <T : Number> T.isBig(): Boolean = this.toDouble() > 10.0

        fun main() {
            println((3).isBig())
            println((42).isBig())
            println(0.5.isBig())
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "false"]);
}

#[test]
fn test_extension_receiver_can_be_nullable_and_differentiated() {
    let out = run_prints(
        r#"
        fun String?.orEmptyTag(): String {
            return if (this == null) "none" else "ok:" + this
        }

        fun main() {
            val left: String? = null
            println("x".orEmptyTag())
            println(left.orEmptyTag())
        }
    "#,
    );
    assert_eq!(out, &["ok:x", "none"]);
}

#[test]
fn test_extension_function_dispatch_uses_static_receiver_type() {
    let out = run_prints(
        r#"
        open class Base
        class Child : Base()

        fun Base.label(): String = "base"
        fun Child.label(): String = "child"

        fun main() {
            val static_base: Base = Child()
            println(static_base.label())
            println(Child().label())
        }
    "#,
    );
    assert_eq!(out, &["base", "child"]);
}

#[test]
fn test_extension_property_on_iterable_reports_head_or_fallback() {
    let out = run_prints(
        r#"
        val <T> Iterable<T>.headOrFallback: String
            get() = if (iterator().hasNext()) iterator().next().toString() else "fallback"

        fun main() {
            println(listOf("a", "b").headOrFallback)
            println(listOf<Int>().headOrFallback)
        }
    "#,
    );
    assert_eq!(out, &["a", "fallback"]);
}

#[test]
fn test_extension_function_on_function_type_calls_receiver_twice() {
    let out = run_prints(
        r#"
        fun (() -> Int).callTwice(): Int {
            return this() + this()
        }

        fun main() {
            var state = 0
            val value = { state += 1; state }
            println(value.callTwice())
            println(value.callTwice())
            println(state)
        }
    "#,
    );
    assert_eq!(out, &["3", "7", "4"]);
}

#[test]
fn test_extension_for_companion_object_acts_like_factory() {
    let out = run_prints(
        r#"
        class Factory private constructor(val value: Int) {
            companion object
        }

        fun Factory.Companion.from(value: Int): Factory = Factory(value)

        fun main() {
            println(Factory.from(7).value)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_extension_function_over_object_expression_receiver() {
    let out = run_prints(
        r#"
        fun StringBuilder.enclosed(): String {
            this.append("]")
            this.insert(0, "[")
            return this.toString()
        }

        fun main() {
            val value = StringBuilder("ok").enclosed()
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["[ok]"]);
}

#[test]
fn test_extension_function_for_nullable_interface_receiver() {
    let out = run_prints(
        r#"
        interface Labelable {
            fun label(): String
        }

        fun Labelable?.labelOrFallback(): String {
            return this?.label() ?: "missing"
        }
        fun main() {
            val missing: Labelable? = null
            val item: Labelable = object : Labelable {
                override fun label(): String = "ok"
            }
            println(item.labelOrFallback())
            println(missing.labelOrFallback())
        }
    "#,
    );
    assert_eq!(out, &["ok", "missing"]);
}
