// vybe-test: kotlin/collections_set/test_set_contains_and_size_with_null_element
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf<String?>(null, "first", null)
            __check((values.size).toString(), "2")
            __check((values.contains(null)).toString(), "true")
            __check((values.contains("missing")).toString(), "false")
        }
