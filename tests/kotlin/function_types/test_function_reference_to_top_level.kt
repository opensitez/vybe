// vybe-test: kotlin/function_types/test_function_reference_to_top_level
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun inc(v: Int): Int = v + 1
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f: (Int) -> Int = ::inc
            __check((f(2)).toString(), "3")
        }
