// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_set_with_comparator
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = sortedSetOf(compareBy<String> { it.length }.thenBy { it }, "bbb", "cc", "a", "ddd")
            __check((values.joinToString(",")).toString(), "a,cc,bbb,ddd")
        }
