// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_with_custom_tie_breaker
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class Item(val first: Int, val second: String)
            val values = listOf(Item(2, "b"), Item(1, "c"), Item(2, "a"))
            val out = values.sortedWith(compareBy<Item> { it.first }.thenBy { it.second })
            __check((out.joinToString(",") { "${'$'}{it.first}${'$'}{it.second}" }).toString(), "1c,2a,2b")
        }
