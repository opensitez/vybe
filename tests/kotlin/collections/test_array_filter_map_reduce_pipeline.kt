// vybe-test: kotlin/collections/test_array_filter_map_reduce_pipeline
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5)
            val result = nums.filter { it % 2 == 1 }
                .map { it * 2 }
                .reduce { acc, value -> acc + value }
            __check((result).toString(), "18")
        }
