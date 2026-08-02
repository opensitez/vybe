// vybe-test: kotlin/scope/test_catch_binding_does_not_clobber_outer_scope
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val message = "root"
            try {
                throw Exception("boom")
            } catch (message: Exception) {
                __check((message.message).toString(), "boom")
            }
            __check(("root").toString(), "root")
        }
