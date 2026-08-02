// vybe-test: kotlin/initialization_order/test_field_initialization_happens_in_declaration_order
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val first = 1
            val second = first + 1
            val third = second + first
            init {
                __check((first).toString(), "1")
                __check((second).toString(), "2")
                __check((third).toString(), "3")
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
