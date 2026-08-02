// vybe-test: kotlin/kotlin_string_split_replace/test_replace_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abcdef"
            val out = StringBuilder(s).replace(1, 3, "ZZ")
            __check((out.toString()).toString(), "aZZdef")
        }
