// vybe-test: kotlin/function_types/test_function_reference_member
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

class A {
            fun show(v: Int): String = "#" + v
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = A()
            val f = a::show
            __check((f(7)).toString(), "#7")
        }
