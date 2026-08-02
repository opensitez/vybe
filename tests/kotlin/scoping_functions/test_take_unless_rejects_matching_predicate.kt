// vybe-test: kotlin/scoping_functions/test_take_unless_rejects_matching_predicate
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7.takeUnless { it == 7 } == null).toString(), "true")
            __check((7.takeUnless { it > 10 } == null).toString(), "false")
        }
