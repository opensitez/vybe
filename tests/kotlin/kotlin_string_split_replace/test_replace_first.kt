// vybe-test: kotlin/kotlin_string_split_replace/test_replace_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "banana"
            __check((s.replaceFirst("ana", "x")).toString(), "bxna")
        }
