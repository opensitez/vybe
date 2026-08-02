// vybe-test: kotlin/kotlin_control_flow_guards/test_when_with_range_and_equality_guards
// origin: languages/kotlin/tests/kotlin/test_kotlin_control_flow_guards.rs

fun status(code: Int): String = when {
            code in 200..299 -> "ok"
            code in 400..499 -> "client"
            code >= 500 -> "server"
            else -> "other"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((status(204)).toString(), "ok")
            __check((status(404)).toString(), "client")
            __check((status(501)).toString(), "server")
            __check((status(101)).toString(), "other")
        }
