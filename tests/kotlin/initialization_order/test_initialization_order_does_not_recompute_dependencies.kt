// vybe-test: kotlin/initialization_order/test_initialization_order_does_not_recompute_dependencies
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var ticks = 0

        class Holder {
            val first = next()
            val second = first + 1

            init {
                __check((first).toString(), "1")
                __check((second).toString(), "2")
            }
        }

        fun next(): Int {
            ticks += 1
            return ticks
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Holder()
            __check((ticks).toString(), "1")
        }
