// vybe-test: kotlin/kotlin_collection_chunking/test_flatten_nested
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = listOf(listOf(1, 2), listOf(3), listOf(4, 5))
            __check((nested.flatten().joinToString(",")).toString(), "1,2,3,4,5")
        }
