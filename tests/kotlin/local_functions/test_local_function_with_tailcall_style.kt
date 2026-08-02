// vybe-test: kotlin/local_functions/test_local_function_with_tailcall_style
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun sum(n: Int, acc: Int = 0): Int {
                return if (n == 0) acc else sum(n - 1, acc + n)
            }
            __check((sum(4)).toString(), "10")
        }
