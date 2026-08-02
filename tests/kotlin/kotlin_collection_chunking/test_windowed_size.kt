// vybe-test: kotlin/kotlin_collection_chunking/test_windowed_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val windows = values.windowed(3)
            __check((windows.size).toString(), "3")
            __check((windows[0].joinToString(",")).toString(), "1,2,3")
            __check((windows[1].joinToString(",")).toString(), "2,3,4")
        }
