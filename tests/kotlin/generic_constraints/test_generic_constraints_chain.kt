// vybe-test: kotlin/generic_constraints/test_generic_constraints_chain
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> chain(v: T): String = "n:" + v.toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((chain(1)).toString(), "n:1")
            __check((chain(2.0)).toString(), "n:2.0")
        }
