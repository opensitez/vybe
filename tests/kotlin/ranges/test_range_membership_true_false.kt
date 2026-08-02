// vybe-test: kotlin/ranges/test_range_membership_true_false
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 4
            val y = 8
            __check((x in 1..6).toString(), "true")
            __check((y in 1 until 6).toString(), "false")
            __check((2 in 6 downTo 1).toString(), "true")
        }
