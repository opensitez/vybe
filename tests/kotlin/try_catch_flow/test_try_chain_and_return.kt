// vybe-test: kotlin/try_catch_flow/test_try_chain_and_return
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun branch(x: Int): Int {
            return try {
                if (x == 0) throw Exception("0")
                x * 2
            } catch (e: Exception) {
                -1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((branch(3)).toString(), "6")
            __check((branch(0)).toString(), "-1")
        }
