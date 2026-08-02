// vybe-test: kotlin/scope/test_scope_after_if_scope_restore
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "start"
            if (true) {
                val value = "if-branch"
                __check((value).toString(), "if-branch")
            }
            __check((value).toString(), "start")
        }
