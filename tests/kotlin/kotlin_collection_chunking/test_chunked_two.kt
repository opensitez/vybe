// vybe-test: kotlin/kotlin_collection_chunking/test_chunked_two
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1..7).toList()
            val chunks = values.chunked(3)
            __check((chunks.size).toString(), "3")
            __check((chunks[0].joinToString(",")).toString(), "1,2,3")
            __check((chunks.last().joinToString(",")).toString(), "7")
        }
