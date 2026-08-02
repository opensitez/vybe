// vybe-test: kotlin/collections_set/test_set_partition_and_counts
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val (small, large) = values.partition { it < 4 }
            __check((small.joinToString(",")).toString(), "1,2,3")
            __check((large.joinToString(",")).toString(), "4,5")
            __check((small.size + large.size).toString(), "5")
        }
