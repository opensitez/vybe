// vybe-test: kotlin/local_functions/test_local_function_with_overloaded_names
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun parseText(value: String): String = "s:" + value
            fun parseText(value: Int): String = "i:" + value
            __check((parseText("x")).toString(), "s:x")
            __check((parseText(8)).toString(), "i:8")
        }
