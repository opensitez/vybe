// vybe-test: kotlin/ranges/test_negative_then_positive_step_is_empty
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val broken = (1..5).step(-2)
            __check((broken.isEmpty()).toString(), "true")
            __check((broken.count()).toString(), "0")
            __check((1 in broken).toString(), "false")
        }
