// vybe-test: kotlin/strings/test_string_length_and_content
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = ""
            val word = "Kotlin"
            __check((empty.length).toString(), "0")
            __check((word.length).toString(), "6")
            __check((empty == "").toString(), "true")
            __check((word == "Kotlin").toString(), "true")
        }
