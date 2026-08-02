// vybe-test: kotlin/extension_functions/test_extension_function_on_any_type
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Any.described(): String = when (this) {
            is Int -> "Int"
            is String -> "String"
            else -> "Any"
        }

        class Item

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.described()).toString(), "Int")
            __check(("x".described()).toString(), "String")
            __check((Item().described()).toString(), "Any")
        }
