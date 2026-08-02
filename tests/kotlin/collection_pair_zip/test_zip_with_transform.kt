// vybe-test: kotlin/collection_pair_zip/test_zip_with_transform
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(2, 4, 6)
            val right = listOf(1, 3, 5)
            val out = left.zip(right) { a, b -> a * b }
            __check((out.joinToString(",")).toString(), "2,12,30")
        }
