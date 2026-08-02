// vybe-test: kotlin/kotlin_collection_chunking/test_distinct_and_retain
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf(1, 1, 2, 3, 3, 4)
            __check((source.distinct().joinToString(",")).toString(), "1,2,3,4")
            __check((source.distinctBy { it % 2 }.joinToString(",")).toString(), "1,2")
        }
