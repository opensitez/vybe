// vybe-test: kotlin/default_arguments/test_default_arguments_boolean_default
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun enabled(flag: Boolean = true): String = if (flag) "yes" else "no"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((enabled()).toString(), "yes")
            __check((enabled(false)).toString(), "no")
        }
