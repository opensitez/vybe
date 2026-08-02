// vybe-test: kotlin/ranges/test_range_first_and_last_properties
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val growing = 1..4
            val declining = 4 downTo 1
            __check((growing.first).toString(), "1")
            __check((growing.last).toString(), "4")
            __check((declining.first).toString(), "4")
            __check((declining.last).toString(), "1")
        }
