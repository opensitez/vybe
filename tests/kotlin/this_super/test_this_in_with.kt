// vybe-test: kotlin/this_super/test_this_in_with
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = "k"
        val y = with(x) { this + this }
        __check((y).toString(), "kk")
    }
