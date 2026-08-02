// vybe-test: kotlin/scope/test_local_function_shadow_parameter
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
            fun inner(value: Int): Int {
                return value + 1
            }
            __check((inner(4)).toString(), "5")
            __check((value).toString(), "3")
        }
