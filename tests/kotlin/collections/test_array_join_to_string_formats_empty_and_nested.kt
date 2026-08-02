// vybe-test: kotlin/collections/test_array_join_to_string_formats_empty_and_nested
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val none = arrayOf<Int>()
            val nested = arrayOf(arrayOf(1), arrayOf(2, 3))
            __check((none.joinToString(",")).toString(), "")
            __check((nested.contentDeepToString()).toString(), "[[1], [2, 3]]")
            __check((arrayOf("a").contentToString()).toString(), "[a]")
        }
