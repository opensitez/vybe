// vybe-test: kotlin/initialization_order/test_primary_constructor_property_initializes_before_init_block
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder(val base: Int) {
            val plus = base + 1
            val label: String

            init {
                label = "v=" + plus.toString()
            }

            fun out(): String = label
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder(3).out()).toString(), "v=4")
        }
