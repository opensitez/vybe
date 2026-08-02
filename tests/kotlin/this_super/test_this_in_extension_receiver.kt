// vybe-test: kotlin/this_super/test_this_in_extension_receiver
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class Box(val n: Int) { fun call() = with(this) { n + 1 } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Box(4).call()).toString(), "5") }
