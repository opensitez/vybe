// vybe-test: kotlin/kotlin_regex_advanced/test_regex_find_all_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val regex = Regex("\\d+")
            val found = regex.findAll("a1 b22 c333")
            __check((found.map { it.value }.joinToString(",")).toString(), "1,22,333")
            __check((found.count()).toString(), "3")
        }
