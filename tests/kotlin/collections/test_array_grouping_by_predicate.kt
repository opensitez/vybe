// vybe-test: kotlin/collections/test_array_grouping_by_predicate
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5, 6)
            val grouped = nums.groupBy { if (it % 2 == 0) "even" else "odd" }
            val even = grouped["even"] ?: arrayOf()
            val odd = grouped["odd"] ?: arrayOf()
            __check((even.joinToString(",")).toString(), "2,4,6")
            __check((odd.joinToString(",")).toString(), "1,3,5")
        }
