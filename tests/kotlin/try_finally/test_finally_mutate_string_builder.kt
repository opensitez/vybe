// vybe-test: kotlin/try_finally/test_finally_mutate_string_builder
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val sb = StringBuilder()
        try {
            sb.append("a")
        } finally {
            sb.append("b")
        }
        __check((sb.toString()).toString(), "ab")
    }
