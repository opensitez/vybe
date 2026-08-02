// vybe-test: kotlin/kotlin_collection_chunking/test_group_by
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val grouped = nums.groupBy { it % 2 == 0 }
            val evens = grouped[true] ?: emptyList<Int>()
            val odds = grouped[false] ?: emptyList<Int>()
            __check((evens.joinToString(",")).toString(), "2,4")
            __check((odds.joinToString(",")).toString(), "1,3")
        }
