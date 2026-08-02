// vybe-test: kotlin/function_types/test_function_type_in_data_flow
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Int = 1
            val f: (Int) -> Int = { it + a }
            val g: (Int) -> Int = { it * a }
            __check((f(2)).toString(), "3")
            __check((g(3)).toString(), "3")
        }
