// vybe-test: kotlin/kotlin_collection_chunking/test_unzip
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = listOf(1 to "a", 2 to "b")
            val (numbers, letters) = pairs.unzip()
            __check((numbers.joinToString(",")).toString(), "1,2")
            __check((letters.joinToString(",")).toString(), "a,b")
        }
