// vybe-test: kotlin/collections/test_two_dimensional_array_access
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grid = arrayOf(
                arrayOf(1, 2),
                arrayOf(3, 4),
            )
            __check((grid[0][1] + grid[1][0]).toString(), "5")
        }
