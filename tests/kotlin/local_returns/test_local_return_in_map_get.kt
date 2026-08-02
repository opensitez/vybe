// vybe-test: kotlin/local_returns/test_local_return_in_map_get
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            val value = map["c"] ?: run {
                if (map.isEmpty()) 0 else 99
            }
            __check((value).toString(), "99")
        }
