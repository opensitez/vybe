// vybe-test: kotlin/random/test_random_seeded_boolean_sequence_is_repeatable
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(19)
            val b = kotlin.random.Random(19)
            __check((a.nextBoolean() == b.nextBoolean()).toString(), "true")
            __check((a.nextBoolean() == b.nextBoolean()).toString(), "true")
            __check((a.nextBoolean() == b.nextBoolean()).toString(), "true")
        }
