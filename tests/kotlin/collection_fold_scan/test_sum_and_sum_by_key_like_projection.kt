// vybe-test: kotlin/collection_fold_scan/test_sum_and_sum_by_key_like_projection
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val rows = listOf(
                Pair("a", 1),
                Pair("b", 2),
                Pair("a", 3)
            )
            __check((rows.sumOf { it.second }).toString(), "6")
            __check((rows.filter { it.first == "a" }.sumOf { it.second }).toString(), "4")
        }
