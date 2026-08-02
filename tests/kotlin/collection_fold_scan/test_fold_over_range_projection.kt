// vybe-test: kotlin/collection_fold_scan/test_fold_over_range_projection
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val projected = (1..4).fold(0) { acc, v -> acc * 10 + v }
            val right = (1..4).foldRight(100) { value, acc -> acc + value }
            __check((projected).toString(), "1234")
            __check((right).toString(), "110")
        }
