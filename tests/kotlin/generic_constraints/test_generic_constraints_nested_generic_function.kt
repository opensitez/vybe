// vybe-test: kotlin/generic_constraints/test_generic_constraints_nested_generic_function
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> outer(v: T): String {
            fun <U : Comparable<U>> inner(a: U, b: U): Boolean = a > b
            return if (inner(v.toString(), "x")) "gt" else "le"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((outer("a")).toString(), "le")
            __check((outer("z")).toString(), "gt")
        }
