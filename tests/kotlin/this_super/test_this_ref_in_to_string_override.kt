// vybe-test: kotlin/this_super/test_this_ref_in_to_string_override
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class A {
        override fun toString() = "A:" + this.javaClass.simpleName
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((A().toString()).toString(), "A:A") }
