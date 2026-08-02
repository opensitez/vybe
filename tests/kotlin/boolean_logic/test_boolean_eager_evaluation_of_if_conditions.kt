// vybe-test: kotlin/boolean_logic/test_boolean_eager_evaluation_of_if_conditions
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            val c = if (a && b) "both" else "not both"
            val d = if (a || b) "some" else "none"
            __check((c).toString(), "not both")
            __check((d).toString(), "some")
        }
