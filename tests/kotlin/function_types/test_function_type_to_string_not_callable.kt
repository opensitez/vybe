// vybe-test: kotlin/function_types/test_function_type_to_string_not_callable
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f: (Int) -> Int = { it + 1 }
            __check((f is Function1<*, *>).toString(), "true")
            __check((f::class.simpleName != null).toString(), "true")
        }
