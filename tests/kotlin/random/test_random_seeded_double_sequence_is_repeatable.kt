// vybe-test: kotlin/random/test_random_seeded_double_sequence_is_repeatable
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(13)
            val b = kotlin.random.Random(13)
            __check((a.nextDouble() == b.nextDouble()).toString(), "true")
            __check((a.nextDouble() == b.nextDouble()).toString(), "true")
        }
