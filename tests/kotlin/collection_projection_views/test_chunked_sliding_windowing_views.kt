// vybe-test: kotlin/collection_projection_views/test_chunked_sliding_windowing_views
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            __check((values.chunked(2).joinToString("|") { it.joinToString(":") }).toString(), "1:2|3:4|5")
            __check((values.windowed(3).joinToString("|") { it.joinToString(":") }).toString(), "1:2:3|2:3:4|3:4:5")
            __check((values.windowed(3, partialWindows = true).joinToString("|") { it.joinToString(":") }).toString(), "1:2:3|2:3:4|3:4:5|4:5|5")
        }
