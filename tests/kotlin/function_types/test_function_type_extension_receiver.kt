// vybe-test: kotlin/function_types/test_function_type_extension_receiver
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val upper: String.() -> String = { uppercase() }
            __check(("a".upper()).toString(), "A")
            val append: String.(String) -> String = { this + it }
            __check(("x".append("y")).toString(), "xy")
        }
