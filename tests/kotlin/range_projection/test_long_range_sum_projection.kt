// vybe-test: kotlin/range_projection/test_long_range_sum_projection
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 2L..6L
            val sum = r.fold(0L) { acc, n -> acc + n }
            __check((sum).toString(), "20")
            __check((r.any { it == 4L }).toString(), "true")
            __check((r.all { it >= 2 }).toString(), "true")
        }
