// vybe-test: kotlin/spread_arguments/test_spread_in_extension_function_call
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun String.tagged(vararg values: String): String = this + values.joinToString(":")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tags = arrayOf("a", "b", "c")
            __check(("base:".tagged(*tags)).toString(), "base:a:b:c")
        }
