// vybe-test: kotlin/extension_functions/test_extension_function_with_default_parameter
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun String.wrap(prefix: String = "x", suffix: String = "!"): String {
            return prefix + this + suffix
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("a".wrap()).toString(), "xa!")
            __check(("a".wrap("z")).toString(), "za!")
            __check(("a".wrap("z", "?")).toString(), "za?")
        }
