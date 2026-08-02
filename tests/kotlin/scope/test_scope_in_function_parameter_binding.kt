// vybe-test: kotlin/scope/test_scope_in_function_parameter_binding
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun square(value: Int): Int {
            return value * value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
            fun emit(value: Int): Int {
                return value + 1
            }
            __check((square(value)).toString(), "9")
            __check((emit(4)).toString(), "5")
            __check((value).toString(), "3")
        }
