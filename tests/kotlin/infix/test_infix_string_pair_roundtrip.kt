// vybe-test: kotlin/infix/test_infix_string_pair_roundtrip
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = "x" to 10
            if (pair.first == "x" && pair.second == 10) {
                __check(("ok").toString(), "ok")
            }
        }
