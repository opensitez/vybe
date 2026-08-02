// vybe-test: kotlin/functions/test_function_no_return_uses_unit
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun sideEffect() {
            __check(("started").toString(), "started")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            sideEffect()
            __check(("done").toString(), "done")
        }
