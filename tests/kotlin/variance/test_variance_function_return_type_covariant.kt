// vybe-test: kotlin/variance/test_variance_function_return_type_covariant
// origin: languages/kotlin/tests/kotlin/test_variance.rs

open class Base
        class Child : Base()
        fun make(): Child = Child()
        fun produce(): Base = make()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b: Base = produce()
            __check((b is Base).toString(), "true")
        }
