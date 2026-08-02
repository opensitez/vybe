// vybe-test: kotlin/collection_fold_scan/test_counting_and_aggregation_shortcuts
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(2, 4, 6, 7, 9)
            __check((values.count { it % 2 == 0 }).toString(), "2")
            __check((values.sumOf { it }).toString(), "28")
            __check((values.average()).toString(), "5.6")
        }
