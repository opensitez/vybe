// vybe-test: kotlin/kotlin_sorting_comparators/test_partition_by_parity
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (even, odd) = listOf(1, 2, 3, 4, 5).partition { it % 2 == 0 }
            __check((even.joinToString(",")).toString(), "2,4")
            __check((odd.joinToString(",")).toString(), "1,3,5")
        }
