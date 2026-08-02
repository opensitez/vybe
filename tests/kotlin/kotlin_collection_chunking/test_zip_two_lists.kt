// vybe-test: kotlin/kotlin_collection_chunking/test_zip_two_lists
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3)
            val b = listOf("a", "b", "c")
            val zipped = a.zip(b)
            __check((zipped.size).toString(), "3")
            __check((zipped[0].first.toString()).toString(), "1")
            __check((zipped[2].second).toString(), "c")
        }
