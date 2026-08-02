// vybe-test: kotlin/runtime_type_queries/test_safe_cast_with_as_question_mark
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Any = "hello"
            val cast1 = a as? String
            val cast2 = a as? Int
            __check((cast1 ?: "none").toString(), "hello")
            __check((cast2?.toString() ?: "none").toString(), "none")
        }
