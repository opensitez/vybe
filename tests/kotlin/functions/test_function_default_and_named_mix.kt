// vybe-test: kotlin/functions/test_function_default_and_named_mix
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun build(prefix: String = "p", suffix: String = "s", count: Int = 1): String {
            return prefix + suffix + count.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((build()).toString(), "ps1")
            __check((build(count = 3)).toString(), "ps3")
            __check((build("a", count = 4, suffix = "b")).toString(), "ab4")
        }
