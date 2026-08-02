// vybe-test: kotlin/function_types/test_function_type_in_try_catch_flow
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun dispatch(v: Int, fn: (Int) -> Int): Int {
            return if (v < 0) 0 else fn(v)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((dispatch(3, { it + 1 })).toString(), "4")
            __check((dispatch(-2, { it + 1 })).toString(), "0")
        }
