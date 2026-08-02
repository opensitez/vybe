// vybe-test: kotlin/function_types/test_function_type_basic_inference
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f: (Int) -> Int = { it + 1 }
            __check((f(2)).toString(), "3")
        }
