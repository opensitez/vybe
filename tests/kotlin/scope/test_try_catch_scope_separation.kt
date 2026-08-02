// vybe-test: kotlin/scope/test_try_catch_scope_separation
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val state = "ok"
            try {
                throw Exception("bad")
            } catch (e: Exception) {
                val state = "caught"
                __check((state).toString(), "caught")
            }
            __check((state).toString(), "ok")
        }
