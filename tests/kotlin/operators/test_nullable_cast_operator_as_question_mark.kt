// vybe-test: kotlin/operators/test_nullable_cast_operator_as_question_mark
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "kotlin"
            val first: String? = value as? String
            val second: Int? = value as? Int
            val third: Any? = null
            val fourth: String? = third as? String
            __check((first).toString(), "kotlin")
            __check((second).toString(), "null")
            __check((fourth).toString(), "null")
        }
