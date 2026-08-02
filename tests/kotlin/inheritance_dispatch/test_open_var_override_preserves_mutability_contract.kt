// vybe-test: kotlin/inheritance_dispatch/test_open_var_override_preserves_mutability_contract
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open var value: Int = 1
        }

        class Child : Base() {
            override var value: Int = 5
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val child = Child()
            child.value += 2
            __check((child.value).toString(), "7")
        }
