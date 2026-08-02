// vybe-test: kotlin/collections_iterables/test_list_distinct_and_union_intersection
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(1, 1, 2, 3, 3)
            val right = listOf(3, 3, 4)
            __check((left.distinct().joinToString(",")).toString(), "1,2,3")
            __check(((left.toSet() intersect right.toSet()).joinToString(",")).toString(), "3")
            __check(((left.toSet() union right.toSet()).joinToString(",")).toString(), "1,2,3,4")
        }
