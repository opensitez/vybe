// vybe-test: kotlin/kotlin_visibility_advanced/test_private_top_level_function_stays_in_file
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

private fun secret(): String = "hidden"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((secret()).toString(), "hidden")
        }
