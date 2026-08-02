// vybe-test: kotlin/this_super/test_this_in_nested_lambda
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K { fun call(): String { val f = { this.toString() }
return f() } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K().call().isNotEmpty()).toString(), "true") }
