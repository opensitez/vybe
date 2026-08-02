// vybe-test: kotlin/scoping_functions/test_let_can_collect_predicate_outcome_count
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var checks = 0
            val value = 3
            val doubled = value.let {
                checks++
                it * 2
            }
            __check((doubled).toString(), "6")
            __check((checks).toString(), "1")
        }
