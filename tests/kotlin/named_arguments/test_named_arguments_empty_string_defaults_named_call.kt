// vybe-test: kotlin/named_arguments/test_named_arguments_empty_string_defaults_named_call
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun wrap(prefix: String = "[", body: String, suffix: String = "]"): String {
            return prefix + body + suffix
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((wrap(body = "x")).toString(), "[x]")
            __check((wrap(prefix = "<", body = "y", suffix = ">")).toString(), "<y>")
        }
