// vybe-test: kotlin/scope/test_inner_scope_does_not_modify_shadowed_outer
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val marker = "root"
            if (marker == "root") {
                val marker = "inner"
                __check((marker).toString(), "inner")
            }
            __check((marker).toString(), "root")
        }
