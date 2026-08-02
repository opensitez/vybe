// vybe-test: kotlin/try_catch_flow/test_try_catch_then_return
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun compute(x: Int): Int {
            try {
                if (x < 0) throw Exception("bad")
                return x + 1
            } catch (e: Exception) {
                return -1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute(1)).toString(), "2")
            __check((compute(-1)).toString(), "-1")
        }
