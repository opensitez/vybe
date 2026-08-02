// vybe-test: kotlin/strings/test_reverse_and_reversed_roundtrip
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "abcd"
            val reversed = word.reversed()
            __check((reversed).toString(), "dcba")
            __check((reversed.reversed()).toString(), "abcd")
        }
