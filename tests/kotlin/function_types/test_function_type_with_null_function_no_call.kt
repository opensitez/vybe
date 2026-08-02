// vybe-test: kotlin/function_types/test_function_type_with_null_function_no_call
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun runMaybe(v: Int, fn: ((Int) -> Int)?): Int {
            return fn?.invoke(v) ?: 0
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((runMaybe(1, null)).toString(), "0")
        }
