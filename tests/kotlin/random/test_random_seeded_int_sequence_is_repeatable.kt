// vybe-test: kotlin/random/test_random_seeded_int_sequence_is_repeatable
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(7)
            val b = kotlin.random.Random(7)
            __check((a.nextInt() == b.nextInt()).toString(), "true")
            __check((a.nextInt() == b.nextInt()).toString(), "true")
        }
