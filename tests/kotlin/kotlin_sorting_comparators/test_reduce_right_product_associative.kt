// vybe-test: kotlin/kotlin_sorting_comparators/test_reduce_right_product_associative
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(2, 3, 4).sortedDescending()
            val product = values.reduce { acc, value -> acc * value }
            __check((values.joinToString(",")).toString(), "4,3,2")
            __check((product).toString(), "24")
        }
