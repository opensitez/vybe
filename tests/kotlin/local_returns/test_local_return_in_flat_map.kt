// vybe-test: kotlin/local_returns/test_local_return_in_flat_map
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = listOf(1, 2, 3).flatMap {
                if (it == 2) return@flatMap listOf<Int>()
                listOf(it)
            }
            __check((data.joinToString("/")).toString(), "1/3")
        }
