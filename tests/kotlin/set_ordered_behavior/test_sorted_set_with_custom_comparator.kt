// vybe-test: kotlin/set_ordered_behavior/test_sorted_set_with_custom_comparator
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.TreeSet<String>(compareByDescending { it })
            values.add("a")
            values.add("c")
            values.add("b")
            __check((values.joinToString(",")).toString(), "c,b,a")
            __check((values.first()).toString(), "c")
            __check((values.last()).toString(), "a")
        }
