// vybe-test: kotlin/inheritance_dispatch/test_field_access_uses_declared_reference_type
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            val value: String = "base"
        }

        class Child : Base() {
            override val value: String = "child"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Child()
            val base: Base = value
            __check((base.value).toString(), "child")
            __check((value.value).toString(), "child")
        }
