// vybe-test: kotlin/collections_iterables/test_list_partition_predicate
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5, 6)
            val (evens, odds) = nums.partition { it % 2 == 0 }
            __check((evens.joinToString(",")).toString(), "2,4,6")
            __check((odds.joinToString(",")).toString(), "1,3,5")
        }
