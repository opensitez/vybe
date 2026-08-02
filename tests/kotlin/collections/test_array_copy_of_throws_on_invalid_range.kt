// vybe-test: kotlin/collections/test_array_copy_of_throws_on_invalid_range
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 2, 3)
            try {
                nums.copyOfRange(3, 2)
            } catch (e: IllegalArgumentException) {
                __check(("bad").toString(), "bad")
            }
        }
