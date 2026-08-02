// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_strings_by_length
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("bbb", "a", "cccc")
            __check((values.sortedBy { it.length }.joinToString(",")).toString(), "a,bbb,cccc")
        }
