// vybe-test: kotlin/initialization_order/test_overridden_property_is_visible_before_derived_fields_init
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base {
            open val label: String = "base"

            init {
                __check((label).toString(), "child")
            }
        }

        class Child : Base() {
            override val label: String = "child"
            val extended = label + ":v"

            init {
                __check((extended).toString(), "child:v")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Child()
        }
