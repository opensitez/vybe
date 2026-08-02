// vybe-test: kotlin/collections_set/test_set_with_nullable_values_distinguishes_null_presence
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: Set<String?> = setOf("a", null, "b", null)
            __check((values.size).toString(), "3")
            __check((values.contains(null)).toString(), "true")
            __check((values.contains("c")).toString(), "false")
        }
