// vybe-test: kotlin/kotlin_type_checks/test_is_as_and_as_question_mark_paths
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<Any?> = listOf("x", 2, null, 3.1)
            val first = values[0] as String
            val second = values[2] as? String
            __check((first).toString(), "x")
            __check((second).toString(), "null")
        }
