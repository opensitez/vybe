// vybe-test: kotlin/set_ordered_behavior/test_tree_set_natural_ordering
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.TreeSet<Int>()
            values.add(5)
            values.add(1)
            values.add(3)
            __check((values.joinToString(",")).toString(), "1,3,5")
            __check((values.first()).toString(), "1")
            __check((values.last()).toString(), "5")
        }
