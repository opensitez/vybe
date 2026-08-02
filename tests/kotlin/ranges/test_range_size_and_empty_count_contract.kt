// vybe-test: kotlin/ranges/test_range_size_and_empty_count_contract
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val full = 1..6
            val down = 6 downTo 1
            val empty = 1..0
            __check((full.count()).toString(), "6")
            __check((down.count()).toString(), "6")
            __check((empty.count()).toString(), "0")
        }
