// vybe-test: kotlin/initialization_order/test_property_init_order_with_override_chain
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base {
            open val base = 1
            init {
                __check((base).toString(), "4")
            }
        }

        class Child : Base() {
            override val base = 4
            init {
                __check((base).toString(), "4")
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
