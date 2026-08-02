// vybe-test: kotlin/local_returns/test_local_return_in_let_binding
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = "x".let {
                if (it.isEmpty()) return@let "empty"
                "val=" + it
            }
            __check((v).toString(), "val=x")
        }
