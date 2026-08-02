// vybe-test: kotlin/initialization_order/test_init_block_can_reference_secondary_defaults
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val value: Int

            init {
                value = 7
                __check(("init").toString(), "init")
            }

            constructor() {
                this()
            }

            fun out(): Int = value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Holder()
            __check((item.out()).toString(), "7")
        }
