// vybe-test: kotlin/initialization_order/test_property_initializers_evaluate_in_declaration_order
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val base = 1
            val multiplied = base * 2
            val summed = multiplied + 1

            init {
                __check((base).toString(), "1")
                __check((multiplied).toString(), "2")
                __check((summed).toString(), "3")
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
