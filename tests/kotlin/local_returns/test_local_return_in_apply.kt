// vybe-test: kotlin/local_returns/test_local_return_in_apply
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = apply("") {
                this += "x"
                return@apply this
            }
            __check((result).toString(), "x")
        }
