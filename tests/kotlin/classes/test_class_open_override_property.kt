// vybe-test: kotlin/classes/test_class_open_override_property
// origin: languages/kotlin/tests/kotlin/test_classes.rs

open class Base {
            open fun label(): String {
                return "base"
            }
        }

        class Child : Base() {
            override fun label(): String {
                return "child"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b: Base = Child()
            __check((b.label()).toString(), "child")
        }
