// vybe-test: kotlin/scope/test_if_else_scoped_binding
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun pick(flag: Boolean): String {
            return if (flag) {
                val value = "yes"
                value
            } else {
                val value = "no"
                value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(true)).toString(), "yes")
            __check((pick(false)).toString(), "no")
        }
