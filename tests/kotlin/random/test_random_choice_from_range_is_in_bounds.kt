// vybe-test: kotlin/random/test_random_choice_from_range_is_in_bounds
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val picked = kotlin.random.Random(73).nextInt(2..9)
            __check((picked in 2..9).toString(), "true")
        }
