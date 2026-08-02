// vybe-test: kotlin/initialization_order/test_init_block_can_read_property_updates
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var factor = 1

        class Holder {
            val value = factor

            init {
                factor = 4
            }

            val adjusted = value * factor

            init {
                __check((value).toString(), "1")
                __check((adjusted).toString(), "4")
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
            __check((factor).toString(), "4")
        }
