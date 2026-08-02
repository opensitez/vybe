// vybe-test: kotlin/function_types/test_function_type_parameter_named
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun apply(value: Int, op: (Int) -> Int): Int {
            return op(value)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(4, { it * 3 })).toString(), "12")
            __check((apply(4) { it + 1 }).toString(), "5")
        }
