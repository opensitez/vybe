// vybe-test: kotlin/this_super/test_this_in_secondary_constructor
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class C(val v: Int) { constructor() : this(3) { __check((this.v).toString(), "3") } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { C() }
