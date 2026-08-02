// vybe-test: kotlin/inheritance_dispatch/test_interface_default_method_can_be_selected_with_super
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Counter {
            fun value(): String = "default"
        }

        class Custom : Counter {
            override fun value(): String = super<Counter>.value() + "-custom"
        }

        class Plain : Counter

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val custom: Counter = Custom()
            val plain: Counter = Plain()
            __check((custom.value()).toString(), "default-custom")
            __check((plain.value()).toString(), "default")
        }
