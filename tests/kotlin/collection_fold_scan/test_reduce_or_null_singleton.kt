// vybe-test: kotlin/collection_fold_scan/test_reduce_or_null_singleton
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = listOf(42)
            __check((one.reduceOrNull { a, b -> a + b }).toString(), "42")
        }
