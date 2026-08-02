// vybe-test: kotlin/nullability/test_safe_cast_failed_returns_null
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
            val base: Base = Base()
            val child = base as? Child
            __check((child == null).toString(), "true")
        }
