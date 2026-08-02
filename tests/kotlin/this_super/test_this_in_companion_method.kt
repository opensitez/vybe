// vybe-test: kotlin/this_super/test_this_in_companion_method
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class C {
        companion object { fun label(): String = "comp" }
        fun out(): String = C.label()
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().out()).toString(), "comp") }
