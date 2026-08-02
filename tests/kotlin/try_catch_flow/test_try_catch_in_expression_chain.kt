// vybe-test: kotlin/try_catch_flow/test_try_catch_in_expression_chain
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun value(x: Int): Int {
            return try {
                if (x == 0) throw IllegalStateException()
                10 / x
            } catch (e: IllegalStateException) {
                -1
            } catch (e: Exception) {
                -2
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((value(0)).toString(), "-1")
            __check((value(2)).toString(), "5")
        }
