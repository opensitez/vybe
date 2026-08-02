// vybe-test: kotlin/initialization_order/test_inherited_property_shadowing_does_not_reorder_init
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base {
            open val value = 2
        }

        class Child : Base() {
            override val value = 7
            val total = value + 1

            init {
                __check((value).toString(), "7")
                __check((total).toString(), "8")
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
