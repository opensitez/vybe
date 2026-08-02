// vybe-test: kotlin/function_types/test_function_type_pass_through_higher_order
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun wrap(f: (Int) -> Int): (Int) -> Int = { n -> f(n) + 1 }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: (Int) -> Int = { it * 2 }
            val wrapped = wrap(base)
            __check((wrapped(3)).toString(), "7")
        }
