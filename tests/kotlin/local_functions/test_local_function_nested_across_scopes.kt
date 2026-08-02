// vybe-test: kotlin/local_functions/test_local_function_nested_across_scopes
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun outer(v: Int): Int {
                fun inner1(x: Int): Int = x + 1
                if (v > 0) {
                    fun inner2(y: Int): Int = inner1(y * 2)
                    return inner2(v)
                }
                return inner1(v)
            }
            __check((outer(3)).toString(), "7")
            __check((outer(0)).toString(), "1")
        }
