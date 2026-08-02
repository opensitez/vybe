// vybe-test: kotlin/scope_shadowing/test_function_parameter_shadows_outer_property
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

val label = "outer"
        fun labelValue(label: String): String {
            return label
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((labelValue("inner")).toString(), "inner")
            __check((label).toString(), "outer")
        }
