// vybe-test: kotlin/collection_fold_scan/test_fold_right_from_end
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c")
            val out = values.foldRight("") { item, acc -> item + acc }
            __check((out).toString(), "abc")
        }
