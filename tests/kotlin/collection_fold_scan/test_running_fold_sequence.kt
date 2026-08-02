// vybe-test: kotlin/collection_fold_scan/test_running_fold_sequence
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3).runningFold(0) { acc, n -> acc + n }.toList()
            __check((out.joinToString(",")).toString(), "0,1,3,6")
        }
