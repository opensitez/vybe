// vybe-test: kotlin/initialization_order/test_secondary_constructor_chain_initializes_once_per_instance
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val value: Int

            init {
                __check(("instance").toString(), "instance")
            }

            constructor() : this(3) {
                __check(("delegated").toString(), "6")
            }

            constructor(seed: Int) {
                value = seed * 2
                __check((value).toString(), "delegated")
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
