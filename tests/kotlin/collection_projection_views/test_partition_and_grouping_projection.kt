// vybe-test: kotlin/collection_projection_views/test_partition_and_grouping_projection
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val (even, odd) = values.partition { it % 2 == 0 }
            __check((even.joinToString(",")).toString(), "2,4")
            __check((odd.joinToString(",")).toString(), "1,3,5")
            val byMod = values.groupBy { it % 2 }
            __check((byMod[0]!!.joinToString(",")).toString(), "2,4")
            __check((byMod[1]!!.joinToString(",")).toString(), "1,3,5")
        }
