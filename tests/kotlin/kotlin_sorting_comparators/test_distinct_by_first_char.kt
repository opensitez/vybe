// vybe-test: kotlin/kotlin_sorting_comparators/test_distinct_by_first_char
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("apple", "ant", "banana", "car")
            __check((values.distinctBy { it[0] }.joinToString(",")).toString(), "apple,banana,car")
        }
