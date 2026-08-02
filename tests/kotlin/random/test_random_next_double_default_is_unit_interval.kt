// vybe-test: kotlin/random/test_random_next_double_default_is_unit_interval
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(31)
            val d = r.nextDouble()
            __check((d >= 0.0).toString(), "true")
            __check((d < 1.0).toString(), "true")
        }
