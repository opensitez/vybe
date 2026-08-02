// vybe-test: kotlin/local_functions/test_local_function_called_from_local_function
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun outer(x: Int): Int {
                fun plusOne(v: Int): Int = v + 1
                fun plusTwo(v: Int): Int = plusOne(v) + 1
                return plusTwo(x)
            }
            __check((outer(3)).toString(), "5")
        }
