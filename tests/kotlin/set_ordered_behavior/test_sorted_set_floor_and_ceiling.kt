// vybe-test: kotlin/set_ordered_behavior/test_sorted_set_floor_and_ceiling
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.TreeSet(listOf(2, 4, 6, 8))
            __check((values.floor(5)).toString(), "4")
            __check((values.ceiling(5)).toString(), "6")
            __check((values.lower(6)).toString(), "4")
            __check((values.higher(6)).toString(), "8")
        }
