// vybe-test: kotlin/random/test_random_next_bytes_repeatable_with_seed
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(53)
            val b = kotlin.random.Random(53)
            val ba = a.nextBytes(4)
            val bb = b.nextBytes(4)
            val same = ba.joinToString(",") == bb.joinToString(",")
            __check((same).toString(), "true")
        }
