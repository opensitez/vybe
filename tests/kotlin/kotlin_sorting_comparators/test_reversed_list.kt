// vybe-test: kotlin/kotlin_sorting_comparators/test_reversed_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4)
            __check((values.reversed().joinToString(",")).toString(), "4,3,2,1")
        }
