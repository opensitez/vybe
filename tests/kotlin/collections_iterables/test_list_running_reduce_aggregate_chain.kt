// vybe-test: kotlin/collections_iterables/test_list_running_reduce_aggregate_chain
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val running = nums.runningReduce { acc, value -> acc + value }
            __check((running.joinToString(",")).toString(), "1,3,6,10")
        }
