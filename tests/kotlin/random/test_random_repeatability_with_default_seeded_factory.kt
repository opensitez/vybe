// vybe-test: kotlin/random/test_random_repeatability_with_default_seeded_factory
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(101)
            val b = kotlin.random.Random(101)
            __check((a.nextInt(0, 1000) == b.nextInt(0, 1000)).toString(), "true")
            __check((a.nextLong(0L, 1000L) == b.nextLong(0L, 1000L)).toString(), "true")
            __check((a.nextBoolean() == b.nextBoolean()).toString(), "true")
        }
