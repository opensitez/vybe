// vybe-test: kotlin/scope_shadowing/test_shadowing_inside_if_branches
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1
            val out = if (value == 1) {
                val value = "one"
                value
            } else {
                "other"
            }
            __check((out).toString(), "one")
            __check((value).toString(), "1")
        }
