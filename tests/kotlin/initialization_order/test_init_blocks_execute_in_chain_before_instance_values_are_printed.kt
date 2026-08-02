// vybe-test: kotlin/initialization_order/test_init_blocks_execute_in_chain_before_instance_values_are_printed
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class LevelOne {
            open val base = "one"
            init {
                __check((base).toString(), "one")
            }
        }

        open class LevelTwo : LevelOne() {
            init {
                __check((base + "-two").toString(), "one-two")
            }
        }

        class LevelThree : LevelTwo() {
            init {
                __check((base + "-three").toString(), "one-three")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            LevelThree()
        }
