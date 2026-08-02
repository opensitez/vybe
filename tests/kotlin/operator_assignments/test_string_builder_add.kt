// vybe-test: kotlin/operator_assignments/test_string_builder_add
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val sb = StringBuilder()
        sb.append("a")
        sb.append("b")
        __check((sb).toString(), "ab")
    }
