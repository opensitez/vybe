// vybe-test: kotlin/try_catch_flow/test_try_finally_and_return
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun value(x: Int): Int {
            try {
                return x * 2
            } finally {
                __check(("done").toString(), "done")
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((value(3)).toString(), "6")
        }
