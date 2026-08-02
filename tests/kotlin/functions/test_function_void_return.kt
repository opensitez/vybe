// vybe-test: kotlin/functions/test_function_void_return
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun doNothing() {
            val a = 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            doNothing()
            __check(("done").toString(), "done")
        }
