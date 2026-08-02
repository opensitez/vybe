// vybe-test: kotlin/initialization_order/test_initialization_of_multiple_properties_order_by_appearance
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val a = 1
            val b = a + c
            val c = 3

            init {
                __check((a).toString(), "1")
                __check((b).toString(), "4")
                __check((c).toString(), "3")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Holder()
        }
