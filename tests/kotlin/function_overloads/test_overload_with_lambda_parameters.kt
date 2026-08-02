// vybe-test: kotlin/function_overloads/test_overload_with_lambda_parameters
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun compute(v: Int, f: (Int) -> Int): Int = f(v)
        fun compute(v: String, f: (String) -> String): String = f(v)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute(3) { it + 1 }).toString(), "4")
            __check((compute("x") { it + "!" }).toString(), "x!")
        }
