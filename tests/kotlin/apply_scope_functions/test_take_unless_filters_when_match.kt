// vybe-test: kotlin/apply_scope_functions/test_take_unless_filters_when_match
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 7.takeUnless { it > 5 }
            val b = 3.takeUnless { it > 5 }
            __check((a).toString(), "null")
            __check((b).toString(), "3")
        }
