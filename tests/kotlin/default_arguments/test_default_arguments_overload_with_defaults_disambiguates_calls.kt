// vybe-test: kotlin/default_arguments/test_default_arguments_overload_with_defaults_disambiguates_calls
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun pick(a: Int, b: Int = 2): Int = a + b
        fun pick(a: String): String = a
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(3)).toString(), "5")
            __check((pick("x")).toString(), "x")
            __check((pick(3, 4)).toString(), "7")
        }
