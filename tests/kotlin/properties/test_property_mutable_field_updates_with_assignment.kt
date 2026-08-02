// vybe-test: kotlin/properties/test_property_mutable_field_updates_with_assignment
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Counter {
            var value: Int = 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            counter.value = 4
            __check((counter.value).toString(), "4")
        }
