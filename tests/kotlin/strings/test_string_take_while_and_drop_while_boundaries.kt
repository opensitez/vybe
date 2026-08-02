// vybe-test: kotlin/strings/test_string_take_while_and_drop_while_boundaries
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "12abc34"
            __check((value.takeWhile { it.isDigit() }).toString(), "12")
            __check((value.dropWhile { it.isDigit() }).toString(), "abc34")
            __check((value.takeLastWhile { it.isDigit() }).toString(), "34")
            __check((value.dropLastWhile { it.isDigit() }).toString(), "12abc")
        }
