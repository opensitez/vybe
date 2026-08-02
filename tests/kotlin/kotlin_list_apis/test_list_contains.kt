// vybe-test: kotlin/kotlin_list_apis/test_list_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf("x", "y", "z")
            __check((list.contains("y")).toString(), "true")
            __check((list.contains("q")).toString(), "false")
        }
