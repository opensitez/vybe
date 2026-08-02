// vybe-test: kotlin/strings_regex/test_regex_find_all_distinct_via_set
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\b\\w+")
            val words = pattern.findAll("a a b c a").map { it.value }.toList().toSet()
            __check((words.size).toString(), "4")
            __check((words.joinToString(",")).toString(), "a,b,c")
        }
