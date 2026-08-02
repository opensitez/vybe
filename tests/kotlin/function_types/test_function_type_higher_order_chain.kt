// vybe-test: kotlin/function_types/test_function_type_higher_order_chain
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun map(v: Int, first: (Int) -> Int, second: (Int) -> Int): Int {
            return second(first(v))
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = map(2, { it + 3 }, { it * 4 })
            __check((out).toString(), "20")
        }
