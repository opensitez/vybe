// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_by_descending_length
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("bbb", "a", "cccc")
            __check((values.sortedByDescending { it.length }.joinToString(",")).toString(), "cccc,bbb,a")
        }
