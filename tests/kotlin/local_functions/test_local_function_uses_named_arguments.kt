// vybe-test: kotlin/local_functions/test_local_function_uses_named_arguments
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun build(a: Int, b: Int = 2, c: Int = 3): Int = a + b + c
            __check((build(a = 1)).toString(), "6")
            __check((build(a = 1, c = 10)).toString(), "13")
        }
