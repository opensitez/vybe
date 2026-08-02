// vybe-test: kotlin/kotlin_collection_chunking/test_chunked_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val transformed = values.chunked(2) { it.sum() }
            __check((transformed.joinToString(",")).toString(), "3,7,5")
        }
