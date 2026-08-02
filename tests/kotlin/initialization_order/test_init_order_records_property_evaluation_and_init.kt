// vybe-test: kotlin/initialization_order/test_init_order_records_property_evaluation_and_init
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var trace = ""
        fun tick(value: String): Int {
            trace += value
            return value.length
        }

        class Holder {
            val first = tick("a")
            val second = tick("b")
            init {
                __check((trace).toString(), "ab")
            }
            val third = first + second
            init {
                __check((third).toString(), "2")
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
