// vybe-test: kotlin/extension_functions/test_extension_receiver_shadowing_from_scope
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun String.show(): String = "global-" + this

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun String.show(): String = "local-" + this
            __check(("x".show()).toString(), "local-x")
        }
