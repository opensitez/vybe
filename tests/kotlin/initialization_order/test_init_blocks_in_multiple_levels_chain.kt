// vybe-test: kotlin/initialization_order/test_init_blocks_in_multiple_levels_chain
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class LevelOne {
            init {
                __check(("one").toString(), "one")
            }
        }

        open class LevelTwo : LevelOne() {
            init {
                __check(("two").toString(), "two")
            }
        }

        class LevelThree : LevelTwo() {
            init {
                __check(("three").toString(), "three")
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
