// vybe-test: kotlin/collections/test_array_content_deep_hashcode_is_stable
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            __check((nested.contentDeepHashCode() > 0).toString(), "true")
            __check((nested.contentDeepToString()).toString(), "[[1, 2], [3, 4]]")
        }
