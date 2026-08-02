// vybe-test: kotlin/ranges/test_range_is_empty_knowledge
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((1..0).isEmpty()).toString(), "true")
            __check(((0 until 0).isEmpty()).toString(), "true")
            __check(((5 downTo 10).isEmpty()).toString(), "true")
        }
