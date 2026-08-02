// vybe-test: kotlin/functions/test_function_nullable_parameter_returns_default
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun describe(input: String?): Int {
            return if (input == null) 0 else input.length
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(null)).toString(), "0")
            __check((describe("abc")).toString(), "3")
        }
