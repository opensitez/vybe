// vybe-test: kotlin/this_super/test_this_in_extension_lambda
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

fun K.tag() = n(this)
class K(val value: Int)
fun n(k: K) = k.value
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K(4).tag()).toString(), "4") }
