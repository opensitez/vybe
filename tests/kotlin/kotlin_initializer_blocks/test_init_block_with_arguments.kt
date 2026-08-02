// vybe-test: kotlin/kotlin_initializer_blocks/test_init_block_with_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_initializer_blocks.rs

class Calc(val base: Int) {
            val offset = base + 1

            init {
                __check(((base * offset).toString()).toString(), "12")
            }

            init {
                __check(((offset - base).toString()).toString(), "1")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Calc(3)
        }
