// vybe-test: kotlin/properties/test_property_jvm_field_backing_name_stable_behavior
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Box {
            var value: Int = 0
            val computed: Int
                get() = value + 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box()
            box.value = 9
            __check((box.computed).toString(), "10")
            box.value = 0
            __check((box.value).toString(), "0")
        }
