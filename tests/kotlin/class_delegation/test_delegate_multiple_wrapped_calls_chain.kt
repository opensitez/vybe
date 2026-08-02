// vybe-test: kotlin/class_delegation/test_delegate_multiple_wrapped_calls_chain
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Op { fun run(x: Int): Int }

        class A : Op { override fun run(x: Int) = x + 1 }
        class B(delegate: Op) : Op by delegate
        class C(delegate: Op) : Op by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = C(B(A()))
            __check((c.run(5)).toString(), "6")
        }
