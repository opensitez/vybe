// vybe-test: kotlin/collections/test_array_diagonal_access
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grid = arrayOf(
                arrayOf(9, 1),
                arrayOf(2, 8),
                arrayOf(3, 4)
            )
            __check((grid[0][0] + grid[1][1] + grid[2][0]).toString(), "20")
        }
