// vybe-test: kotlin/properties/test_property_in_local_class_scope
// origin: languages/kotlin/tests/kotlin/test_properties.rs

fun makeTag(prefix: String): String {
            class Box {
                val label: String = prefix + "-box"
            }
            return Box().label
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((makeTag("new")).toString(), "new-box")
        }
