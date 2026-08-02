// vybe-test: kotlin/collection_projection_views/test_distinct_and_union_projection
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 2, 3)
            val b = listOf(2, 3, 4)
            val uniq = a.distinct()
            val union = a.union(b)
            __check((uniq.joinToString(",")).toString(), "1,2,3")
            __check((union.joinToString(",")).toString(), "1,2,2,3,4")
        }
