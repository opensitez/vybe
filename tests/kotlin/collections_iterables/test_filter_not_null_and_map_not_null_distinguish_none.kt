// vybe-test: kotlin/collections_iterables/test_filter_not_null_and_map_not_null_distinguish_none
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf(1, null, 2, null, 3)
            __check((items.filterNotNull().joinToString(",")).toString(), "1,2,3")
            val transformed = items.mapNotNull { it?.plus(10) }
            __check((transformed.joinToString(",")).toString(), "11,12,13")
        }
