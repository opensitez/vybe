// vybe-test: kotlin/named_arguments/test_named_arguments_empty_call_path
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun pick(a: String, b: String = "x", c: String = "y"): String = a + b + c
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(a = "1", b = "2", c = "3")).toString(), "123")
        }
