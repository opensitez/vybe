// vybe-test: kotlin/generic_constraints/test_generic_constraints_reifiable_like_behavior
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> show(v: T): String = "" + v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((show(1)).toString(), "1")
            __check((show(1.2)).toString(), "1.2")
        }
