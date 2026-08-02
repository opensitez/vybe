// vybe-test: kotlin/strings_regex/test_regex_split_to_sequence
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("-")
            val parts = pattern.splitToSequence("x-y-z").toList()
            __check((parts.size).toString(), "3")
            __check((parts.joinToString(",")).toString(), "x,y,z")
        }
