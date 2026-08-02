// vybe-test: kotlin/when_expressions/test_when_with_boolean_subject
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun label(flag: Boolean): String {
            return when (flag) {
                true -> "on"
                false -> "off"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(true)).toString(), "on")
            __check((label(false)).toString(), "off")
        }
