// vybe-test: kotlin/collections/test_array_first_last_and_singleton_helpers
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(8, 6, 7)
            __check((nums.first()).toString(), "8")
            __check((nums.last()).toString(), "7")
            __check((arrayOf("solo").single()).toString(), "solo")
            __check((nums.take(2).joinToString(",")).toString(), "8,6")
            __check((nums.drop(1).joinToString(",")).toString(), "6,7")
        }
