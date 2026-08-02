// vybe-test: kotlin/collections_iterables/test_list_chunked_and_windowed
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.chunked(2).joinToString("|") { it.joinToString("-") }).toString(), "1-2|3-4|5")
            __check((nums.windowed(2).joinToString("|") { it.joinToString("-") }).toString(), "1-2|2-3|3-4|4-5")
        }
