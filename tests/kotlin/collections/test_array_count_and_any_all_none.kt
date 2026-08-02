// vybe-test: kotlin/collections/test_array_count_and_any_all_none
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 3, 5, 6)
            __check((nums.count { it > 3 }).toString(), "2")
            __check((nums.any { it == 6 }).toString(), "true")
            __check((nums.all { it > 0 }).toString(), "true")
            __check((nums.none { it < 0 }).toString(), "true")
        }
