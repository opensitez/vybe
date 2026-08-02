// vybe-test: kotlin/scope_shadowing/test_shadowing_for_same_file_scope_not_persisting
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val marker = "x"
            fun first() { val marker = "y"
__check((marker).toString(), "y") }
            fun second() { __check((marker).toString(), "x") }
            first()
            second()
        }
