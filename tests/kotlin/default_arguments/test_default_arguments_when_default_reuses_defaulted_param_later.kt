// vybe-test: kotlin/default_arguments/test_default_arguments_when_default_reuses_defaulted_param_later
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun score(base: Int = 2, factor: Int = base * 2): Int = factor
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score()).toString(), "4")
            __check((score(5)).toString(), "10")
            __check((score(1, 9)).toString(), "9")
        }
