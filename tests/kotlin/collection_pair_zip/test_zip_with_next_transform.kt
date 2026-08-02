// vybe-test: kotlin/collection_pair_zip/test_zip_with_next_transform
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val mapped = nums.zipWithNext { a, b -> a + b }
            __check((mapped.joinToString(",")).toString(), "3,5,7")
        }
