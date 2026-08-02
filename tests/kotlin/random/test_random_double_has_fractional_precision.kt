// vybe-test: kotlin/random/test_random_double_has_fractional_precision
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(109)
            val value = r.nextDouble()
            __check((value.toString().contains(".")).toString(), "true")
        }
