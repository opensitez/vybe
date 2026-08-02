// vybe-test: kotlin/function_overloads/test_overload_same_name_across_extensions
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

class Host
        fun Host.label(): String = "host"
        fun Host.label(prefix: String): String = prefix + ":host"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Host()
            __check((h.label()).toString(), "host")
            __check((h.label("x")).toString(), "x:host")
        }
