// vybe-test: kotlin/collections_iterables/test_list_running_fold_accumulates_prefixes
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val prefixes = nums.runningFold(0) { acc, value -> acc + value }
            __check((prefixes.joinToString(",")).toString(), "0,1,3,6,10")
        }
