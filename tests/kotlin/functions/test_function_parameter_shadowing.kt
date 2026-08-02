// vybe-test: kotlin/functions/test_function_parameter_shadowing
// origin: languages/kotlin/tests/kotlin/test_functions.rs

val x = 100

        fun testShadow(x: Int) {
            __check((x).toString(), "5")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            testShadow(5)
            __check((x).toString(), "100")
        }
