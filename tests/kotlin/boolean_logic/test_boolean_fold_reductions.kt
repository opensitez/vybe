// vybe-test: kotlin/boolean_logic/test_boolean_fold_reductions
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = booleanArrayOf(true, false, true, true)
            __check((values.fold(true) { acc, value -> acc && value }).toString(), "false")
            __check((values.fold(false) { acc, value -> acc || value }).toString(), "true")
            __check((values.reduce { acc, value -> acc && value }).toString(), "false")
            __check((values.reduce { acc, value -> acc || value }).toString(), "true")
        }
