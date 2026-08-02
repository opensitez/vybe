// vybe-test: kotlin/collections/test_array_find_last_if_empty
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(2, 4, 6)
            __check((nums.find { it > 5 }).toString(), "6")
            __check((nums.findLast { it < 0 } ?: -1).toString(), "-1")
            __check((nums.firstOrNull { it == 10 } ?: "missing").toString(), "missing")
        }
