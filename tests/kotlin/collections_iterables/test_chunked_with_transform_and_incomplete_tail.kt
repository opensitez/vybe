// vybe-test: kotlin/collections_iterables/test_chunked_with_transform_and_incomplete_tail
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = (1..5).toList()
            __check((nums.chunked(2).joinToString("|") { it.joinToString("-") }).toString(), "1-2|3-4|5")
            __check((nums.chunked(3) { it.sum() }.joinToString(",")).toString(), "6,9")
        }
