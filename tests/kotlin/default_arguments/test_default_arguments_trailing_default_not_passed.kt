// vybe-test: kotlin/default_arguments/test_default_arguments_trailing_default_not_passed
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun combine(prefix: String, postfix: String = "X", center: String = "Y"): String {
            return prefix + center + postfix
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((combine("a")).toString(), "aYX")
            __check((combine("a", "B")).toString(), "aYB")
            __check((combine("a", center = "C", postfix = "D")).toString(), "aCD")
        }
