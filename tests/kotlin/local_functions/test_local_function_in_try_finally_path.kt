// vybe-test: kotlin/local_functions/test_local_function_in_try_finally_path
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var executed = false
            fun body(v: Int): Int {
                executed = true
                return v
            }
            val value = try {
                body(9)
            } finally {
                __check((if (executed) "ok" else "missing").toString(), "ok")
            }
            __check((value).toString(), "9")
        }
