// vybe-test: kotlin/collections/test_explicit_typed_array
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums: Array<Int> = arrayOf(5, 6, 7)
            __check((nums.size).toString(), "3")
            __check((nums[2]).toString(), "7")
        }
