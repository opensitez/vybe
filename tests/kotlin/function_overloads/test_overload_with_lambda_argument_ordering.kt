// vybe-test: kotlin/function_overloads/test_overload_with_lambda_argument_ordering
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun exec(v: Int, f: () -> String): String = "i:" + f()
        fun exec(v: String, f: () -> String): String = "s:" + f()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((exec(1) { "x" }).toString(), "i:x")
            __check((exec("y") { "z" }).toString(), "s:z")
        }
