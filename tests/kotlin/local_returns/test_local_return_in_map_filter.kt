// vybe-test: kotlin/local_returns/test_local_return_in_map_filter
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val filtered = listOf(1, 2, 3).filter {
                if (it % 2 == 0) return@filter false
                true
            }
            __check((filtered.joinToString()).toString(), "1, 3")
        }
