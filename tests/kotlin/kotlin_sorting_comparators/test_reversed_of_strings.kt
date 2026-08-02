// vybe-test: kotlin/kotlin_sorting_comparators/test_reversed_of_strings
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c")
            __check((values.reversed()).toString(), "[c, b, a]")
        }
