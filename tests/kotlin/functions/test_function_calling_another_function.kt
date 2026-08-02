// vybe-test: kotlin/functions/test_function_calling_another_function
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun doubleVal(x: Int): Int = x * 2
        fun tripleVal(x: Int): Int = doubleVal(x) + x

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((tripleVal(4)).toString(), "12")
        }
