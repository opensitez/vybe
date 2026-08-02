// vybe-test: kotlin/kotlin_string_split_replace/test_replace_without_match
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abc"
            __check((s.replace("z", "q")).toString(), "abc")
        }
