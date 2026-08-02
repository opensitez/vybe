// vybe-test: kotlin/collection_pair_zip/test_zip_empty_inputs
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf<Int>().zip(listOf("a")).isEmpty()).toString(), "true")
            __check((listOf<Int>().zipWithNext().isEmpty()).toString(), "true")
        }
