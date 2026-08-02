// vybe-test: kotlin/random/test_random_next_int_with_lower_and_upper_bound
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(5)
            val first = r.nextInt(-4, 7)
            val second = r.nextInt(-4, 7)
            val third = r.nextInt(-4, 7)
            __check((first in -4..6).toString(), "true")
            __check((second in -4..6).toString(), "true")
            __check((third in -4..6).toString(), "true")
        }
