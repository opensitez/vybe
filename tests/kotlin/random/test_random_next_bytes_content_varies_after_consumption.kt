// vybe-test: kotlin/random/test_random_next_bytes_content_varies_after_consumption
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(59)
            val first = r.nextBytes(2)
            val second = r.nextBytes(2)
            __check((first.joinToString(",") == second.joinToString(",")).toString(), "false")
        }
