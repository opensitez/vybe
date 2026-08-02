// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_drop_take_consistency
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(9, 4, 7, 1).sorted()
            __check((values.drop(2).joinToString(",")).toString(), "7,9")
            __check((values.take(2).joinToString(",")).toString(), "1,4")
        }
