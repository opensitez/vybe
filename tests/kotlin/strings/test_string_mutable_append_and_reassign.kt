// vybe-test: kotlin/strings/test_string_mutable_append_and_reassign
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = "Hello"
            value += ", "
            value += "World"
            __check((value).toString(), "Hello, World")
            __check((value.length).toString(), "12")
        }
