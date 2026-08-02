// vybe-test: kotlin/scope/test_scope_across_returned_function_body
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun makeGreeter(prefix: String): (Int) -> String {
            val suffix = "!"
            return { value ->
                val body = prefix + value.toString()
                body + suffix
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val greet = makeGreeter("x")
            __check((greet(1)).toString(), "x1!")
            __check((greet(2)).toString(), "x2!")
        }
