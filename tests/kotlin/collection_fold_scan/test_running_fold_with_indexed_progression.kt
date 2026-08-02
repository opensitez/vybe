// vybe-test: kotlin/collection_fold_scan/test_running_fold_with_indexed_progression
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parts = listOf(1, 1, 2, 3).runningFoldIndexed(0) { index, acc, item ->
                acc + item + index
            }
            __check((parts.joinToString(",")).toString(), "0,1,3,6,10")
        }
