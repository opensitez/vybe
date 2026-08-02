// vybe-test: kotlin/collection_projection_views/test_map_not_null_projection
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: List<Int?> = listOf(1, null, 3, null, 5)
            __check((source.filterNotNull().joinToString(",")).toString(), "1,3,5")
            __check((source.mapNotNull { it?.plus(10) }.joinToString(",")).toString(), "11,13,15")
        }
