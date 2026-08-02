// vybe-test: kotlin/function_types/test_function_type_with_two_args
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun combine(a: Int, b: Int, fn: (Int, Int) -> Int): Int {
            return fn(a, b)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((combine(2, 4, { x, y -> x + y })).toString(), "6")
            __check((combine(2, 4, Int::plus)).toString(), "6")
        }
