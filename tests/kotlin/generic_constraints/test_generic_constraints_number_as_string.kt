// vybe-test: kotlin/generic_constraints/test_generic_constraints_number_as_string
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> asString(v: T): String = v.toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asString(10)).toString(), "10")
            __check((asString(10.8)).toString(), "10.8")
        }
