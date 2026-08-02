// vybe-test: kotlin/block_expressions/test_block_with_try_catch
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = try {
            run {
                throw IllegalArgumentException()
            }
        } catch (e: Exception) {
            11
        }
        __check((x).toString(), "11")
    }
