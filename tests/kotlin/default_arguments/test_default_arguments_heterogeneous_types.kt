// vybe-test: kotlin/default_arguments/test_default_arguments_heterogeneous_types
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun combine(a: Int = 1, b: String = "x", c: Boolean = false): String {
            return a.toString() + b + (if (c) "Y" else "N")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((combine()).toString(), "1xN")
            __check((combine(2, c = true)).toString(), "2xY")
            __check((combine(b = "z", a = 3)).toString(), "3zN")
        }
