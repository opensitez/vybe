// vybe-test: kotlin/collection_fold_scan/test_reduce_right_on_strings
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("x", "y", "z")
            val out = values.reduceRight { item, acc -> item + "," + acc }
            __check((out).toString(), "x,y,z")
        }
