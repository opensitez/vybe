// vybe-test: kotlin/invoke_operator/test_invoke_operator_overload_heterogeneous_types
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class C {
            operator fun invoke(v: Int): String = "i$v"
            operator fun invoke(v: Double): String = "d$v"
            operator fun invoke(v: Boolean): String = if (v) "on" else "off"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = C()
            __check((c(1)).toString(), "i1")
            __check((c(2.5)).toString(), "d2.5")
            __check((c(false)).toString(), "off")
        }
