// vybe-test: kotlin/scope/test_scope_nested_function_mutates_outer_var_after_shadowing
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 5

            fun bump(delta: Int) {
                fun total(): Int {
                    return value + delta
                }
                value = total()
            }

            bump(3)
            __check((value).toString(), "8")

            val value = 10
            fun useShadowed(): Int {
                return value + 1
            }
            __check((useShadowed()).toString(), "11")
        }
