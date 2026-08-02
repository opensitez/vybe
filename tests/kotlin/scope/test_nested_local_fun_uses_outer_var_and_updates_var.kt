// vybe-test: kotlin/scope/test_nested_local_fun_uses_outer_var_and_updates_var
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 1
            fun inc(step: Int) {
                value += step
            }
            inc(2)
            val value = 9
            __check((value).toString(), "9")
            inc(3)
            __check((value).toString(), "9")
            __check((value == 9).toString(), "true")
        }
