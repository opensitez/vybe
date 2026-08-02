// vybe-test: kotlin/kotlin_collection_chunking/test_windowed_step
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5, 6)
            val windows = values.windowed(2, step = 2)
            __check((windows.size).toString(), "3")
            __check((windows.joinToString("|") { it.joinToString(",") }).toString(), "1,2|3,4|5,6")
        }
