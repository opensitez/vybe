// vybe-test: kotlin/kotlin_sorting_comparators/test_scan_accumulation_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3).sorted()
            val accumulated = values.runningReduce { acc, n -> acc + n }
            __check((accumulated.joinToString(",")).toString(), "1,3,6")
        }
