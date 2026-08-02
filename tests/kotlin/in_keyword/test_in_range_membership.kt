// vybe-test: kotlin/in_keyword/test_in_range_membership
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3 in 1..5).toString(), "true")
            __check((7 in 1..5).toString(), "false")
            __check((5 !in 1 until 5).toString(), "true")
        }
