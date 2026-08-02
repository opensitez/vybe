// vybe-test: kotlin/random/test_random_next_double_with_bounds
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(41)
            val d = r.nextDouble(-1.0, 2.0)
            __check((d >= -1.0).toString(), "true")
            __check((d < 2.0).toString(), "true")
        }
