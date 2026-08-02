// vybe-test: kotlin/local_functions/test_local_function_compiles_with_boolean_logic
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun isEven(v: Int): Boolean = (v % 2 == 0)
            fun describe(v: Int): String = if (isEven(v)) "even" else "odd"
            __check((describe(4)).toString(), "even")
            __check((describe(5)).toString(), "odd")
        }
