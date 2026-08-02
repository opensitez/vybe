// vybe-test: kotlin/kotlin_type_parameter_bounds/test_where_clause_in_generic_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

fun <T> toText(value: T): String where T : Any {
            return value.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((toText("ok")).toString(), "ok")
            __check((toText(5)).toString(), "5")
        }
