// vybe-test: kotlin/when_expressions/test_when_with_nested_subject_binding_in_same_when
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun decode(value: Any): String {
            return when (value) {
                is Int -> {
                    val doubled = value * 2
                    when {
                        doubled > 10 -> "int-big"
                        else -> "int-small"
                    }
                }
                is String -> {
                    val head = value.firstOrNull() ?: '?'
                    when (head) {
                        in 'a'..'m' -> "string-low"
                        in 'n'..'z' -> "string-high"
                        else -> "string-other"
                    }
                }
                else -> "none"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((decode(7)).toString(), "int-small")
            __check((decode(4)).toString(), "int-small")
            __check((decode("beta")).toString(), "string-low")
            __check((decode("zeta")).toString(), "string-high")
            __check((decode("@")).toString(), "string-other")
            __check((decode(3.0)).toString(), "none")
        }
