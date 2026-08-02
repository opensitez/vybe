// vybe-test: kotlin/random/test_random_next_double_with_upper_bound
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(37)
            val d = r.nextDouble(1.5)
            __check((d >= 0.0).toString(), "true")
            __check((d < 1.5).toString(), "true")
        }
