// vybe-test: kotlin/this_super/test_this_nested_class_calls
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class A { val v = 3
    inner class B { fun v() = this@A.v }
}
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((A().B().v()).toString(), "3") }
