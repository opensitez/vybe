// vybe-test: kotlin/ranges/test_char_range_count_and_reverse_view
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val span = 'b'..'f'
            __check((span.count()).toString(), "5")
            val reversed = span.reversed()
            __check((reversed.count()).toString(), "5")
            __check((reversed.first()).toString(), "f")
            __check((reversed.last()).toString(), "b")
        }
