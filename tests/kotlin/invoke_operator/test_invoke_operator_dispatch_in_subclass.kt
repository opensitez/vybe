// vybe-test: kotlin/invoke_operator/test_invoke_operator_dispatch_in_subclass
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

open class A {
            open operator fun invoke(v: String): String = "A: " + v
        }
        class B : A() {
            override operator fun invoke(v: String): String = "B: " + v
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: A = B()
            __check((a("x")).toString(), "B: x")
        }
