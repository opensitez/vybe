// vybe-test: kotlin/random/test_random_moving_state_advances_sequence
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(103)
            val first = r.nextInt()
            val second = r.nextInt()
            val third = kotlin.random.Random(103).nextInt()
            __check((first == second).toString(), "false")
            __check((first == third).toString(), "true")
        }
