// vybe-test: kotlin/type_aliases/test_typealias_for_nullable_function_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias OptionalText = (() -> String)?

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: OptionalText = null
            __check(((value == null)).toString(), "true")
            __check(((value?.invoke() ?: "empty")).toString(), "empty")
        }
