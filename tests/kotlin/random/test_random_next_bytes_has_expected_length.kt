// vybe-test: kotlin/random/test_random_next_bytes_has_expected_length
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(47)
            __check((r.nextBytes(3).size).toString(), "3")
            __check((r.nextBytes(5).size).toString(), "5")
        }
