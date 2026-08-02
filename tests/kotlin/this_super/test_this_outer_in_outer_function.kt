// vybe-test: kotlin/this_super/test_this_outer_in_outer_function
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class Outer {
        val t = "outer"
        inner class Inner { fun t() = this@Outer.t }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Outer().Inner().t()).toString(), "outer") }
