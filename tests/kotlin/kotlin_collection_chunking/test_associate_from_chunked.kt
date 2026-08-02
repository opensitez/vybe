// vybe-test: kotlin/kotlin_collection_chunking/test_associate_from_chunked
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = listOf(1, 2, 3)
            val map = pairs.chunked(1).associate { chunk ->
                val key = chunk[0]
                key to "v$key"
            }
            __check((map.size).toString(), "3")
            __check((map[2]).toString(), "v2")
        }
