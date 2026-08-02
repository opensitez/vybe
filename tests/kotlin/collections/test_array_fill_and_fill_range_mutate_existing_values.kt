// vybe-test: kotlin/collections/test_array_fill_and_fill_range_mutate_existing_values
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = IntArray(5) { it }
            nums.fill(9)
            __check((nums.joinToString(",")).toString(), "9,9,9,9,9")
            val src = intArrayOf(1, 2, 3, 4, 5)
            src.fill(7, 1, 4)
            __check((src.joinToString(",")).toString(), "1,7,7,7,5")
        }
