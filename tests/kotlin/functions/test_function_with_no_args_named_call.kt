// vybe-test: kotlin/functions/test_function_with_no_args_named_call
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun getStatus(prefix: String = "OK", code: Int = 200): String {
            return prefix + code.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((getStatus()).toString(), "OK200")
            __check((getStatus(code = 404)).toString(), "OK404")
        }
