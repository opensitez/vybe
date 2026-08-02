// vybe-test: kotlin/type_aliases/test_typealias_for_function_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Reducer = (Int, Int) -> Int

        fun combine(value: Int, other: Int, op: Reducer): Int {
            return op(value, other)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = combine(4, 5, { a, b -> a + b })
            __check((result).toString(), "9")
        }
