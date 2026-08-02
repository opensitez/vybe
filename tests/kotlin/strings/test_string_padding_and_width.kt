// vybe-test: kotlin/strings/test_string_padding_and_width
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "7"
            __check((value.padStart(3, "0")).toString(), "007")
            __check((value.padEnd(4, "_")).toString(), "7___")
            __check(("abc".padStart(5, "x")).toString(), "xxabc")
        }
