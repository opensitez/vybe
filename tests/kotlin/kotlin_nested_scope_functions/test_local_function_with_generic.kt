// vybe-test: kotlin/kotlin_nested_scope_functions/test_local_function_with_generic
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun <T> firstOf(items: List<T>): T = items[0]
            __check((firstOf(listOf("a", "b"))).toString(), "a")
            __check((firstOf(listOf(1, 2, 3))).toString(), "1")
        }
