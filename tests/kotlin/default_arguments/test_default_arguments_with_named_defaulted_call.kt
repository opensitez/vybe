// vybe-test: kotlin/default_arguments/test_default_arguments_with_named_defaulted_call
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun emit(prefix: String, text: String = "ok", suffix: String = ""): String = prefix + text + suffix
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((emit("<", suffix = ">")).toString(), "<ok>")
        }
