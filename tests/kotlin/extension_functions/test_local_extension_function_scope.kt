// vybe-test: kotlin/extension_functions/test_local_extension_function_scope
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun String.shout(): String = this.uppercase()
            fun use(value: String): String = value.shout()
            __check((use("go")).toString(), "GO")
        }
