// vybe-test: kotlin/functions/test_function_throws_and_finally_blocks_do_not_swallow_return_value
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun risky(value: Int): Int {
            try {
                if (value < 0) {
                    throw Exception("bad")
                }
                return value * 2
            } finally {
                println("final")
            }
        }

        fun main() {
            println(risky(4))
            try {
                risky(-1)
            } catch (e: Exception) {
                println("caught")
            }
        }

