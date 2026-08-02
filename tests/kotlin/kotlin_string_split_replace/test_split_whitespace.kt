// vybe-test: kotlin/kotlin_string_split_replace/test_split_whitespace
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "one two  three"
            val parts = s.split(" ")
            __check((parts.size).toString(), "3")
            __check((parts[0]).toString(), "one")
        }
