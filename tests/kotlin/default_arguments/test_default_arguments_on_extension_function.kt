// vybe-test: kotlin/default_arguments/test_default_arguments_on_extension_function
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun String.wrap(prefix: String = "<", suffix: String = ">"): String {
            return prefix + this + suffix
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("a".wrap()).toString(), "<a>")
            __check(("b".wrap(prefix = "[")).toString(), "[b]")
        }
