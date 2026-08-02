// vybe-test: kotlin/collections_set/test_set_with_string_predicate
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf("alpha", "beta", "gamma")
            __check((values.count { it.length >= 5 }).toString(), "2")
            __check((values.joinToString("|")).toString(), "alpha|beta|gamma")
            __check((values.any { it.startsWith("a") }).toString(), "true")
        }
