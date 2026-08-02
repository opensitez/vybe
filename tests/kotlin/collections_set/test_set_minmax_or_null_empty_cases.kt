// vybe-test: kotlin/collections_set/test_set_minmax_or_null_empty_cases
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((setOf<Int>().minOrNull() ?: -1).toString(), "-1")
            __check((setOf<Int>().maxOrNull() ?: -1).toString(), "-1")
        }
