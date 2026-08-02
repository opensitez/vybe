// vybe-test: kotlin/recursion/test_recursion_nested_map
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun flatten(items: List<List<Int>>, row: Int = 0): List<Int> {
            if (row >= items.size) return listOf()
            return items[row] + flatten(items, row + 1)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = flatten(listOf(listOf(1), listOf(2, 3)))
            __check((values.joinToString(".")).toString(), "1.2.3")
        }
