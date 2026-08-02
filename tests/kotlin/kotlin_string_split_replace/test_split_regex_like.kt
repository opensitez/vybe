// vybe-test: kotlin/kotlin_string_split_replace/test_split_regex_like
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "1;2;3"
            val parts = s.split(";")
            __check((parts[0] + parts[2]).toString(), "13")
        }
