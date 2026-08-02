// vybe-test: kotlin/inheritance_dispatch/test_virtual_property_dispatch_in_inheritance_chain
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open val value: Int = 1
            open fun total(): Int = value + 1
        }

        class Child : Base() {
            override val value: Int = 3
            override fun total(): Int = value + 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Base = Child()
            __check((item.value).toString(), "3")
            __check((item.total()).toString(), "5")
        }
