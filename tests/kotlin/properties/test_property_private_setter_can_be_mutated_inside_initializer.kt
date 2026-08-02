// vybe-test: kotlin/properties/test_property_private_setter_can_be_mutated_inside_initializer
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Counter {
            var value: Int = 0
                private set

            init {
                value = 7
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter().value).toString(), "7")
        }
