// vybe-test: kotlin/generic_constraints/test_generic_constraints_any_to_string
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> render(v: T): String {
            return v.toString()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(true)).toString(), "true")
        }
