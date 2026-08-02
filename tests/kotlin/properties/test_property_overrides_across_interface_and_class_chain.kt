// vybe-test: kotlin/properties/test_property_overrides_across_interface_and_class_chain
// origin: languages/kotlin/tests/kotlin/test_properties.rs

interface Named {
            val name: String
        }

        open class Animal : Named {
            override val name: String = "animal"
        }

        class Dog : Animal() {
            override val name: String = "dog"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val named: Named = Dog()
            __check((named.name).toString(), "dog")
        }
