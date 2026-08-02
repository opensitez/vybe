// vybe-test: kotlin/kotlin_receiver_extension_ops/test_simple_string_extension
// origin: languages/kotlin/tests/kotlin/test_kotlin_receiver_extension_ops.rs

fun String.wrap(prefix: String): String = prefix + "-" + this

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("core".wrap("pre")).toString(), "pre-core")
        }
