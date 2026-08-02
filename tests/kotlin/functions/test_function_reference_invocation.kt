// vybe-test: kotlin/functions/test_function_reference_invocation
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun double(x: Int): Int = x * 2

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f: (Int) -> Int = ::double
            __check((f(7)).toString(), "14")
            __check((::double(3)).toString(), "6")
        }
