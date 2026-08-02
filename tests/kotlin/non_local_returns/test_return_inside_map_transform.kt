// vybe-test: kotlin/non_local_returns/test_return_inside_map_transform
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun total(values: List<Int>): Int {
            values.map {
                if (it > 10) return 100
                it
            }
            return -1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((total(listOf(3, 8, 12))).toString(), "100")
        }
