// vybe-test: kotlin/recursion/test_recursion_sum_nested_arrays
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun nestedSum(items: List<List<Int>>, row: Int = 0): Int {
            return if (row >= items.size) 0 else items[row].sum() + nestedSum(items, row + 1)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((nestedSum(listOf(listOf(1, 2), listOf(3, 4)))).toString(), "10")
        }
