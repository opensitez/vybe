// vybe-test: kotlin/set_ordered_behavior/test_tree_set_descending_set_view
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.TreeSet<Int>()
            values.addAll(listOf(4, 1, 3, 2))
            val down = values.descendingSet()
            __check((down.joinToString(",")).toString(), "4,3,2,1")
            __check((down.first()).toString(), "4")
            __check((down.last()).toString(), "1")
        }
