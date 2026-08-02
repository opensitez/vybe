// vybe-test: kotlin/kotlin_collection_chunking/test_windowed_join
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = "abccba"
            val windows = values.windowed(2, partialWindows = false)
            __check((windows.size).toString(), "5")
            __check((windows.joinToString("|")).toString(), "ab|bc|cc|cb|ba")
        }
