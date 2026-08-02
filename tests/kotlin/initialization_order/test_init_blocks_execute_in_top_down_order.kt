// vybe-test: kotlin/initialization_order/test_init_blocks_execute_in_top_down_order
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Base {
            init {
                __check(("base").toString(), "base")
            }
        }

        class Leaf : Base() {
            init {
                __check(("leaf").toString(), "leaf")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Leaf()
        }
