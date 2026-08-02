// vybe-test: kotlin/initialization_order/test_initialization_logs_in_nested_property_chain
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val first = 1
            val second = first + one()
            val third = second + 1

            init {
                __check((third).toString(), "4")
            }
        }

        fun one(): Int = 2

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Holder()
        }
