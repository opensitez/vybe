// vybe-test: kotlin/collections_iterables/test_distinct_by_projection_keeps_first_per_group
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("alpha", "alpine", "beta", "breeze", "bravo")
            __check((words.distinctBy { it[0] }.joinToString(",")).toString(), "alpha,beta")
        }
