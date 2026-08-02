// vybe-test: kotlin/kotlin_range_edge_conditions/test_integer_range_to_list_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((100..102).toList().size).toString(), "3")
            __check(((100 downTo 100).toList().size).toString(), "1")
        }
