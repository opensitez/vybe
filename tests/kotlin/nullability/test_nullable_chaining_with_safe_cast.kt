// vybe-test: kotlin/nullability/test_nullable_chaining_with_safe_cast
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

open class Base
        class Child : Base()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Base? = Child()
            val child = base as? Child
            if (child != null) {
                __check(("ok").toString(), "ok")
            }
        }
