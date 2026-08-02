// vybe-test: kotlin/collections_set/test_set_zip_with_index_like_build
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf("a", "b", "c")
            val zip = values.zipWithNext()
            __check((zip.size).toString(), "2")
            __check((zip[0].first).toString(), "a")
            __check((zip[0].second).toString(), "b")
        }
