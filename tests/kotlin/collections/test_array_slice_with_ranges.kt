// vybe-test: kotlin/collections/test_array_slice_with_ranges
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(10, 20, 30, 40, 50)
            val part = nums.sliceArray(1..3)
            val tail = nums.copyOfRange(3, 5)
            __check((part.joinToString(",")).toString(), "20,30,40")
            __check((tail.joinToString(",")).toString(), "40,50")
        }
