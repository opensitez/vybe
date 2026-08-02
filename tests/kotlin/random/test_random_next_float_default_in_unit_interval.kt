// vybe-test: kotlin/random/test_random_next_float_default_in_unit_interval
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(43)
            val f = r.nextFloat()
            __check((f >= 0.0f).toString(), "true")
            __check((f < 1.0f).toString(), "true")
        }
