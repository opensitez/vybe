// vybe-test: kotlin/type_aliases/test_typealias_for_comparator_reference
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias IntSort = Comparator<Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(4, 1, 3, 2)
            val order: IntSort = Comparator { left, right -> right - left }
            __check((values.sortedWith(order).joinToString("-")).toString(), "4-3-2-1")
        }
