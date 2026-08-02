// vybe-test: kotlin/collections_set/test_set_sum_and_avg_via_fold
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4)
            val total = values.fold(0) { acc, value -> acc + value }
            val avg = total / values.size
            __check((total).toString(), "10")
            __check((avg).toString(), "2")
        }
