// vybe-test: kotlin/random/test_random_next_int_with_upper_bound_uses_bound
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = kotlin.random.Random(3)
            __check((r.nextInt(10) >= 0).toString(), "true")
            __check((r.nextInt(10) < 10).toString(), "true")
            __check((r.nextInt(10) >= 0 && r.nextInt(10) < 10).toString(), "true")
        }
