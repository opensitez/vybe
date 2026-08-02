// vybe-test: kotlin/when_expressions/test_when_with_local_type_checks_and_smart_casts
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun convert(value: Any): String {
            return when (value) {
                is Int -> "i=" + value.toString()
                is Long -> "l=" + value.toString()
                is Double -> "d=" + value.toString()
                else -> "x"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((convert(3)).toString(), "i=3")
            __check((convert(4L)).toString(), "l=4")
            __check((convert(1.5)).toString(), "d=1.5")
            __check((convert("x")).toString(), "x")
        }
