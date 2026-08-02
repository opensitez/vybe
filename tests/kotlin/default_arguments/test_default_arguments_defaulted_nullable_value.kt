// vybe-test: kotlin/default_arguments/test_default_arguments_defaulted_nullable_value
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun pick(v: String?, fallback: String = "d"): String = v ?: fallback
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(null)).toString(), "d")
            __check((pick("x")).toString(), "x")
            __check((pick("", fallback = "z")).toString(), "z")
        }
