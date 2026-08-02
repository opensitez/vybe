// vybe-test: kotlin/random/test_random_next_long_with_upper_bound
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(29)
            val v = r.nextLong(100L)
            __check((v >= 0).toString(), "true")
            __check((v < 100).toString(), "true")
            __check((v in 0L..99L).toString(), "true")
        }
