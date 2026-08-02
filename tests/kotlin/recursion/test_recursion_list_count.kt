// vybe-test: kotlin/recursion/test_recursion_list_count
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun count(items: List<Int>): Int = if (items.isEmpty()) 0 else 1 + count(items.drop(1))
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((count(listOf(1, 2, 3))).toString(), "3")
        }
