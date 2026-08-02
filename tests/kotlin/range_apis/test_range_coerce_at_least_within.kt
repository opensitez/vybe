// vybe-test: kotlin/range_apis/test_range_coerce_at_least_within
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..5
            val value = r.coerceIn(3)
            __check((value).toString(), "3")
            val under = r.coerceIn(0)
            __check((under).toString(), "1")
            val over = r.coerceIn(7)
            __check((over).toString(), "5")
        }
