// vybe-test: kotlin/this_super/test_this_in_constructor
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K(val x: Int) { init { __check((this.x).toString(), "3") } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { K(3) }
