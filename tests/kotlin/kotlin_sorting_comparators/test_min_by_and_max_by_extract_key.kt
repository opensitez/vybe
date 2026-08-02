// vybe-test: kotlin/kotlin_sorting_comparators/test_min_by_and_max_by_extract_key
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class Item(val id: Int, val label: String)
            val values = listOf(Item(3, "z"), Item(1, "x"), Item(2, "y"))
            __check((values.minByOrNull { it.id }?.label).toString(), "x")
            __check((values.maxByOrNull { it.id }?.label).toString(), "z")
        }
