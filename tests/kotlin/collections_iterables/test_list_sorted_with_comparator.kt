// vybe-test: kotlin/collections_iterables/test_list_sorted_with_comparator
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf("delta", "a", "charlie", "bb")
            val sorted = nums.sortedWith(compareBy { it.length })
            __check((sorted.joinToString(",")).toString(), "a,bb,charlie,delta")
            val reverse = nums.sortedWith(compareByDescending { it.length })
            __check((reverse.joinToString(",")).toString(), "delta,charlie,bb,a")
        }
