// vybe-test: kotlin/ranges/test_range_equality_and_string_representation
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val rangeA = 1..3
            val rangeB = 1..3
            val rangeC = 1..4
            __check((rangeA == rangeB).toString(), "true")
            __check((rangeA == rangeC).toString(), "false")
            __check((rangeA.toString()).toString(), "1..3")
        }
