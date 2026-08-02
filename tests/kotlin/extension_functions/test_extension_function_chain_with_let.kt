// vybe-test: kotlin/extension_functions/test_extension_function_chain_with_let
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun String.repeatPrefix(prefix: String, count: Int): String = prefix.repeat(count) + this

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "k"
                .repeatPrefix("a", 3)
                .repeatPrefix("b", 2)
            __check((value).toString(), "bbaaaak")
        }
