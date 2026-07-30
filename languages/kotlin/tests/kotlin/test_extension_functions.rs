use crate::helpers::run_prints;

#[test]
fn test_extension_function_on_primitive() {
    let out = run_prints(r#"
        fun Int.incremented(): Int = this + 1

        fun main() {
            println(3.incremented())
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_extension_function_with_multiple_parameters() {
    let out = run_prints(r#"
        fun Int.add(value: Int, scale: Int): Int = (this + value) * scale

        fun main() {
            println(2.add(3, 4))
        }
    "#);
    assert_eq!(out, &["20"]);
}

#[test]
fn test_extension_function_on_class_instance() {
    let out = run_prints(r#"
        class Box(val value: Int)

        fun Box.labeled(prefix: String): String = prefix + ":" + value

        fun main() {
            println(Box(7).labeled("v"))
        }
    "#);
    assert_eq!(out, &["v:7"]);
}

#[test]
fn test_extension_property_getter() {
    let out = run_prints(r#"
        class Point(val x: Int, val y: Int)

        val Point.sum: Int
            get() = x + y

        fun main() {
            println(Point(2, 5).sum)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_extension_property_with_setter_like_behavior() {
    let out = run_prints(r#"
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
    "#);
    assert_eq!(out, &["5", "10"]);
}

#[test]
fn test_extension_function_for_nullable_receiver() {
    let out = run_prints(r#"
        fun Int?.orZero(): Int = this ?: 0

        fun main() {
            val value: Int? = null
            val second: Int? = 7
            println(value.orZero())
            println(second.orZero())
        }
    "#);
    assert_eq!(out, &["0", "7"]);
}

#[test]
fn test_local_extension_function_scope() {
    let out = run_prints(r#"
        fun main() {
            fun String.shout(): String = this.uppercase()
            fun use(value: String): String = value.shout()
            println(use("go"))
        }
    "#);
    assert_eq!(out, &["GO"]);
}

#[test]
fn test_overload_resolution_between_extension_and_member() {
    let out = run_prints(r#"
        class Box {
            fun value(): Int = 1
        }

        fun Box.value(): Int = 4

        fun main() {
            println(Box().value())
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_generic_extension_transform() {
    let out = run_prints(r#"
        fun <T> List<T>.wrapCount(): String = "count=" + this.size

        fun main() {
            println(listOf(1, 2, 3).wrapCount())
            println(listOf("a").wrapCount())
        }
    "#);
    assert_eq!(out, &["count=3", "count=1"]);
}

#[test]
fn test_extension_on_generic_with_bounds() {
    let out = run_prints(r#"
        fun <T : Number> T.asIntText(): Int = this.toInt()

        fun main() {
            println(4.9.asIntText())
            println(7.asIntText())
        }
    "#);
    assert_eq!(out, &["4", "7"]);
}

#[test]
fn test_extension_function_chain_with_let() {
    let out = run_prints(r#"
        fun String.repeatPrefix(prefix: String, count: Int): String = prefix.repeat(count) + this

        fun main() {
            val value = "k"
                .repeatPrefix("a", 3)
                .repeatPrefix("b", 2)
            println(value)
        }
    "#);
    assert_eq!(out, &["bbaaaak"]);
}
