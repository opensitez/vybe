// vybe-test: kotlin/type_aliases/test_typealias_with_array_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias IntArrayLike = Array<Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: IntArrayLike = arrayOf(1, 2)
            __check((values.size).toString(), "2")
            __check((values[0] + values[1]).toString(), "3")
        }
