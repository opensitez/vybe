// vybe-test: kotlin/generic_constraints/test_generic_constraints_defaulted_callable
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

class Box<T>(private val v: T)
        fun <T> valueOrEmpty(v: T?): String = v?.toString() ?: "empty"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: String? = null
            val b: String? = "x"
            __check((valueOrEmpty(a)).toString(), "empty")
            __check((valueOrEmpty(b)).toString(), "x")
        }
