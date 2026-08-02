// vybe-test: kotlin/initialization_order/test_init_of_derived_uses_base_initialized_state
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base {
            val base = 10
            init { __check((base).toString(), "10") }
        }

        class Child : Base() {
            val child = base + 1
            init { __check((child).toString(), "11") }
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
