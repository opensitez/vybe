// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_with_then_by_secondary
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class Item(val first: Int, val second: Int)
            val values = listOf(Item(1, 9), Item(1, 2), Item(2, 5))
            val out = values.sortedWith(compareBy<Item> { it.first }.thenByDescending { it.second })
            __check((out.joinToString(",") { "${'$'}{it.first}-${'$'}{it.second}" }).toString(), "1-9,1-2,2-5")
        }
