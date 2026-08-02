// vybe-test: kotlin/strings/test_string_padstart_and_padend_short_circuit_when_width_too_small
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("abcdef".padStart(3, "x")).toString(), "abcdef")
            __check(("abcdef".padEnd(2, "x")).toString(), "abcdef")
            __check(("a".padStart(3, ".")).toString(), "..a")
            __check(("a".padEnd(3, ".")).toString(), "a..")
        }
