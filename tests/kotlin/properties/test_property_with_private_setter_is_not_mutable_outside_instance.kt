// vybe-test: kotlin/properties/test_property_with_private_setter_is_not_mutable_outside_instance
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Counter {
            var value: Int = 1
                private set

            fun add(next: Int) {
                value += next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            counter.add(4)
            __check((counter.value).toString(), "5")
        }
