// vybe-test: kotlin/random/test_random_list_shuffled_with_seed_is_repeatable
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(67).shuffled(listOf(1, 2, 3, 4))
            val b = kotlin.random.Random(67).shuffled(listOf(1, 2, 3, 4))
            __check((a == b).toString(), "true")
            __check((a.size).toString(), "4")
        }
