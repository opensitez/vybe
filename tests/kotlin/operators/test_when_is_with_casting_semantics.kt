// vybe-test: kotlin/operators/test_when_is_with_casting_semantics
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun describe(value: Any?): String {
            return when (value) {
                is String -> "str:" + value.length
                is Int -> "int"
                null -> "nil"
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("kotlin")).toString(), "str:6")
            __check((describe(7)).toString(), "int")
            __check((describe(3.14)).toString(), "other")
            __check((describe(null)).toString(), "nil")
        }
