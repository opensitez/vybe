// vybe-test: kotlin/infix/test_infix_to_construction
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val pair = 10 to 20
__check((pair.first).toString(), "10")
__check((pair.second).toString(), "20") }
