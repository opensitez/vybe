// vybe-test: kotlin/collection_fold_scan/test_reduce_vs_reduce_or_null
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3, 4)
            val r = a.reduce { acc, value -> acc + value }
            __check((r).toString(), "10")
            val b = emptyList<Int>()
            __check((b.reduceOrNull() ?: "empty").toString(), "empty")
        }
