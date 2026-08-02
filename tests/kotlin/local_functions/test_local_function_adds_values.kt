// vybe-test: kotlin/local_functions/test_local_function_adds_values
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun add(a: Int, b: Int): Int = a + b
            __check((add(2, 3)).toString(), "5")
        }
