// vybe-test: kotlin/range_apis/test_long_range_map_to_ints
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1L..3L
            val sum = r.map { it.toInt() }.sum()
            __check((sum).toString(), "6")
        }
