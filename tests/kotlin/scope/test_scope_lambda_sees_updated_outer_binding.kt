// vybe-test: kotlin/scope/test_scope_lambda_sees_updated_outer_binding
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var prefix = "before"
            val format = { value: String ->
                prefix + ":" + value
            }

            __check((format("one")).toString(), "before:one")

            prefix = "after"
            __check((format("two")).toString(), "after:two")
        }
