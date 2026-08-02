// vybe-test: kotlin/properties/test_property_val_reference_still_mutable_object
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder {
            val values = mutableListOf(1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder()
            holder.values.add(2)
            __check((holder.values.size).toString(), "2")
        }
