// vybe-test: kotlin/member_references/test_bound_member_function_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Greeter(val name: String) {
            fun hello(prefix: String) = "$prefix$name"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val g = Greeter("k")
            val hi = g::hello
            __check((hi("x")).toString(), "xk")
        }
