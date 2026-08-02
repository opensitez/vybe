// vybe-test: kotlin/collection_fold_scan/test_fold_right_indexed
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(4, 5, 6)
            val out = values.foldRightIndexed("") { index, value, acc ->
                acc + value.toString() + "#" + index.toString() + ";"
            }
            __check((out).toString(), "6#2;5#1;4#0;")
        }
