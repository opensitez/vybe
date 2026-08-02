// vybe-test: kotlin/local_functions/test_nested_local_function_call_chain
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun outer(x: Int): Int {
                fun inner(y: Int): Int = y + 1
                return inner(x) * 2
            }
            __check((outer(7)).toString(), "16")
        }
