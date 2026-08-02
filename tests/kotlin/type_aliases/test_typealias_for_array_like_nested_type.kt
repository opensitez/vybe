// vybe-test: kotlin/type_aliases/test_typealias_for_array_like_nested_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Matrix = Array<IntArray>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grid: Matrix = arrayOf(intArrayOf(1, 2), intArrayOf(3, 4))
            __check((grid.size).toString(), "2")
            __check((grid[1][0]).toString(), "3")
        }
