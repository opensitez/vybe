// vybe-test: kotlin/random/test_random_list_shuffle_with_seed_is_repeatable
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(61)
            val b = kotlin.random.Random(61)
            val sourceA = mutableListOf(1, 2, 3, 4, 5)
            val sourceB = mutableListOf(1, 2, 3, 4, 5)
            sourceA.shuffle(a)
            sourceB.shuffle(b)
            __check((sourceA == sourceB).toString(), "true")
            __check((sourceA.size).toString(), "5")
        }
