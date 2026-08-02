// vybe-test: kotlin/properties/test_top_level_property_and_function_scope_interaction
// origin: languages/kotlin/tests/kotlin/test_properties.rs

var prefix = "A"

        fun scoped(next: String): String {
            return prefix + ":" + next
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((scoped("ok")).toString(), "A:ok")
            prefix = "B"
            __check((scoped("ok")).toString(), "B:ok")
        }
