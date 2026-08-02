// vybe-test: kotlin/random/test_random_next_long_with_bounds_stays_in_range
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(23)
            val value = r.nextLong(-10L, 20L)
            __check((value >= -10).toString(), "true")
            __check((value < 20).toString(), "true")
            __check((value in -10L..19L).toString(), "true")
        }
