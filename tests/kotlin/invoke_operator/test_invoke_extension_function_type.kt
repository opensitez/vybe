// vybe-test: kotlin/invoke_operator/test_invoke_extension_function_type
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Host {
            operator fun String.invoke(v: String): String = this + v
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val host = Host()
            with(host) {
                __check(("a".invoke("b")).toString(), "ab")
            }
        }
