// vybe-test: kotlin/local_functions/test_local_function_uses_outer_scope
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 10
            fun scale(x: Int): Int = x * base
            __check((scale(4)).toString(), "40")
        }
