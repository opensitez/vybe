// vybe-test: kotlin/random/test_random_choose_with_reproducible_seed
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = listOf("a", "b", "c", "d")
            val a = src.random(kotlin.random.Random(71))
            val b = src.random(kotlin.random.Random(71))
            __check((a == b).toString(), "true")
        }
