// vybe-test: kotlin/named_arguments/test_named_arguments_lambda_argument_name_in_callable_reference
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun apply(label: String, fn: (String) -> String = { it }): String {
            return fn(label)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(label = "x", fn = { "v:" + it })).toString(), "v:x")
            __check((apply(label = "y")).toString(), "y")
        }
