// vybe-test: kotlin/type_aliases/test_typealias_receiver_function_invocation_from_aliased_builder
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias StringBuilderRecipe = StringBuilder.() -> Unit

        fun formatValue(value: String): String {
            val recipe: StringBuilderRecipe = {
                append(value)
                append("-done")
            }
            val target = StringBuilder()
            target.recipe()
            return target.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((formatValue("ok")).toString(), "ok-done")
        }
