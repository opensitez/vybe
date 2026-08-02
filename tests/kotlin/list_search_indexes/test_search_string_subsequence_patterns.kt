// vybe-test: kotlin/list_search_indexes/test_search_string_subsequence_patterns
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lines = listOf("alpha", "beta", "gamma", "alphabet")
            __check((lines.indexOf("beta")).toString(), "1")
            __check((lines.indexOfFirst { it.startsWith("alp") }).toString(), "0")
            __check((lines.indexOfLast { it.endsWith("ta") }).toString(), "2")
        }
