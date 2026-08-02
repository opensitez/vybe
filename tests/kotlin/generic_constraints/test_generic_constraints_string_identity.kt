// vybe-test: kotlin/generic_constraints/test_generic_constraints_string_identity
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> identityString(v: T): String = v.toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((identityString(7)).toString(), "7")
            __check((identityString("x")).toString(), "x")
        }
