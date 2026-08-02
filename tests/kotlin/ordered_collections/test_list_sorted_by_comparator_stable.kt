// vybe-test: kotlin/ordered_collections/test_list_sorted_by_comparator_stable
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

data class Pair(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(Pair(1, 2), Pair(1, 1), Pair(0, 3))
            val sorted = list.sortedWith(compareBy<Pair> { it.left }.thenBy { it.right })
            __check((sorted.map { "${'$'}{it.left}:${'$'}{it.right}" }.joinToString("|")).toString(), "0:3|1:1|1:2")
        }
