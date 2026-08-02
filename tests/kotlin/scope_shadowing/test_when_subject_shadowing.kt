// vybe-test: kotlin/scope_shadowing/test_when_subject_shadowing
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val token = "root"
            val result = when (token) {
                "root" -> {
                    val token = 100
                    token
                }
                else -> token.length
            }
            __check((result).toString(), "100")
            __check((token).toString(), "root")
        }
