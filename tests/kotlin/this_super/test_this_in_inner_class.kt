// vybe-test: kotlin/this_super/test_this_in_inner_class
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class Outer {
        val tag = "outer"
        inner class Inner { fun outerTag() = this@Outer.tag }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Outer().Inner().outerTag()).toString(), "outer") }
