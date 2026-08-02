// vybe-test: kotlin/properties/test_property_companion_object_shared_state
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Factory {
            companion object {
                var created: Int = 0
            }

            init {
                Factory.created += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Factory()
            Factory()
            __check((Factory.created).toString(), "2")
        }
