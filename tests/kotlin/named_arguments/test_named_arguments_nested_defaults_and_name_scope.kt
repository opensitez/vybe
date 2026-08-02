// vybe-test: kotlin/named_arguments/test_named_arguments_nested_defaults_and_name_scope
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun outer(tag: String, a: Int = 1, b: Int = 2): Int {
            return a + b
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((outer("t", b = 6)).toString(), "7")
            __check((outer(tag = "u", a = 2, b = 8)).toString(), "10")
        }
