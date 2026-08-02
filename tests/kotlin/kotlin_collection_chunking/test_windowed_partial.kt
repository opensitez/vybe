// vybe-test: kotlin/kotlin_collection_chunking/test_windowed_partial
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3)
            val windows = values.windowed(2, partialWindows = true)
            __check((windows.size).toString(), "2")
            __check((windows[2].joinToString(",")).toString(), "3")
        }
