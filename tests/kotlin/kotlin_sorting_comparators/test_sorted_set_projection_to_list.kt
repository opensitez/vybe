// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_set_projection_to_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = sortedSetOf("b", "a", "c")
            __check((values.toList().joinToString(",")).toString(), "a,b,c")
            __check((values.toMutableList()[0]).toString(), "a")
        }
