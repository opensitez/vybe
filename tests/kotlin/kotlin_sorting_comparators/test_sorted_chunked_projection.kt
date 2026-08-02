// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_chunked_projection
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(5, 2, 9, 1, 8, 3).sorted()
            val chunked = values.chunked(3).joinToString("|") { it.joinToString(",") }
            __check((chunked).toString(), "1,2,3|5,8,9")
        }
