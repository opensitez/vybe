// vybe-test: kotlin/local_returns/test_local_return_in_list_any_match
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val hasLarge = listOf(1, 2, 8, 9).any {
                if (it >= 8) return@any true
                false
            }
            __check((hasLarge).toString(), "true")
        }
