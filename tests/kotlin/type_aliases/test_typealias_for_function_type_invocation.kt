// vybe-test: kotlin/type_aliases/test_typealias_for_function_type_invocation
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Join = (String, String) -> String

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val joiner: Join = { left, right -> left + right }
            __check((joiner("a", "b")).toString(), "ab")
        }
