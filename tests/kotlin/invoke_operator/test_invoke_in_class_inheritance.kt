// vybe-test: kotlin/invoke_operator/test_invoke_in_class_inheritance
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

open class Base {
            operator fun invoke(v: Int): Int = v
        }
        class Child : Base()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child()(9)).toString(), "9")
        }
